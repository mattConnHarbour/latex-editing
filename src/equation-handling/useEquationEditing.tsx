import { useCallback, useEffect, useMemo, useRef, useState, type ReactNode, type RefObject } from 'react';
import type { SuperDocRef, ToolbarConfig } from '@superdoc/react';
import type { BrowserDocumentApi, SelectionTarget } from 'superdoc/ui';
import EquationEditor from './EquationEditor';
import {
  deleteEquation,
  equationOoxmlFromLatex,
  insertEquation,
  loadEquationDocument,
  mathParagraphOoxmlFromLatex,
  replaceEquation,
  type EquationDocument,
  type EquationInsertionPoint,
  type EquationReference,
} from './docx-equation-transformer';
import { captureScrollSnapshot, getEquationEditorPosition, restoreScrollSnapshot, type ScrollSnapshot } from './equation-positioning';
import 'katex/dist/katex.min.css';
import './equation-editor.css';

type ActiveEquation = {
  latex: string;
  position: { left: number; top: number };
} & (
  | { mode: 'edit'; equationNumber: number; reference: EquationReference; selectionTarget: SelectionTarget | null }
  | { mode: 'insert'; insertionPoint: EquationInsertionPoint; selectionTarget: SelectionTarget }
);

export type UseEquationEditingOptions = {
  editorRef: RefObject<SuperDocRef | null>;
  initialDocument: string | Blob;
  onError?: (error: unknown) => void;
};

export type EquationEditingController = {
  documentSource: string | Blob;
  documentVersion: number;
  editor: ReactNode;
  editorHostClassName: string;
  editorHostRef: RefObject<HTMLDivElement | null>;
  onReady: () => void;
  ready: boolean;
  status: string;
  toolbarConfig: ToolbarConfig;
};

type ToolbarButton = Extract<NonNullable<ToolbarConfig['customItems']>[number], { type: 'button' }>;
type ToolbarSelectContext = Parameters<NonNullable<ToolbarButton['onSelect']>>[0];
type SelectionTargetHost = {
  getHandles(): {
    editing?: {
      selectionTargets?: {
        applyAndWait?(target: SelectionTarget): Promise<unknown>;
      };
    } | null;
  };
};
const OOXML_NAMESPACES = {
  m: 'http://schemas.openxmlformats.org/officeDocument/2006/math',
  w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
};

function isRangeSelection(target: SelectionTarget | null): target is SelectionTarget {
  if (!target) return false;
  return JSON.stringify(target.start) !== JSON.stringify(target.end);
}

function equationRangeFromCaret(target: SelectionTarget | null, reference: EquationReference): SelectionTarget | null {
  if (!target || target.start.kind !== 'text' || target.end.kind !== 'text') return null;
  if (target.start.blockId !== target.end.blockId) return null;
  return {
    ...target,
    start: { ...target.start, offset: reference.startOffset },
    end: { ...target.end, offset: reference.endOffset },
  };
}

function equationRangeFromBlock(blockId: string, reference: EquationReference): SelectionTarget {
  const story = { kind: 'story' as const, storyType: 'body' as const };
  return {
    kind: 'selection',
    story,
    start: { kind: 'text', blockId, offset: reference.startOffset, story },
    end: { kind: 'text', blockId, offset: reference.endOffset, story },
  };
}

function successfulReceipt(receipt: unknown): receipt is { success: true } {
  return Boolean(receipt && typeof receipt === 'object' && 'success' in receipt && receipt.success === true);
}

function bytesFromBase64(value: string) {
  const binary = atob(value);
  return Uint8Array.from(binary, (character) => character.charCodeAt(0));
}

function equationReferenceFromClick(math: Element, document: EquationDocument) {
  const fragment = math.closest<HTMLElement>('.superdoc-fragment[data-source-anchor]');
  if (!fragment) return null;
  try {
    const anchor = JSON.parse(fragment.dataset.sourceAnchor ?? '{}') as { sourceRef?: { xpathLikePath?: string } };
    const match = /body\/w:p\[ordinal=(\d+)\]/.exec(anchor.sourceRef?.xpathLikePath ?? '');
    if (!match) return null;
    const equationIndexInParagraph = Array.from<Element>(fragment.querySelectorAll('math')).indexOf(math);
    const equationNumber = document.equations.findIndex(
      (equation) => equation.paragraphOrdinal === Number(match[1]) && equation.equationIndexInParagraph === equationIndexInParagraph,
    );
    const reference = document.equations[equationNumber];
    const blockId = fragment.dataset.sourceNodeId ?? fragment.dataset.blockId;
    return reference && blockId ? { blockId, equationNumber: equationNumber + 1, reference } : null;
  } catch {
    return null;
  }
}

function insertionPointFromSelection(host: HTMLElement, blockId: string, offset: number) {
  const fragment = Array.from(host.querySelectorAll<HTMLElement>('.superdoc-fragment[data-source-anchor]'))
    .find((candidate) => candidate.dataset.sourceNodeId === blockId || candidate.dataset.blockId === blockId);
  if (!fragment) return null;
  try {
    const anchor = JSON.parse(fragment.dataset.sourceAnchor ?? '{}') as {
      sourceRef?: { partUri?: string; xpathLikePath?: string };
    };
    if (anchor.sourceRef?.partUri !== '/word/document.xml') return null;
    const match = /body\/w:p\[ordinal=(\d+)\]/.exec(anchor.sourceRef.xpathLikePath ?? '');
    return match
      ? { fragment, insertionPoint: { offset, paragraphOrdinal: Number(match[1]) } }
      : null;
  } catch {
    return null;
  }
}

export function useEquationEditing({ editorRef, initialDocument, onError }: UseEquationEditingOptions): EquationEditingController {
  const editorHostRef = useRef<HTMLDivElement>(null);
  const activeMathRef = useRef<Element | null>(null);
  const pendingScrollRef = useRef<ScrollSnapshot | null>(null);
  const onErrorRef = useRef(onError);
  const equationDocumentRef = useRef<EquationDocument | null>(null);
  const readyRef = useRef(false);
  const [documentSource, setDocumentSource] = useState<string | Blob>(initialDocument);
  const [documentVersion, setDocumentVersion] = useState(0);
  const [equationDocument, setEquationDocument] = useState<EquationDocument | null>(null);
  const [activeEquation, setActiveEquation] = useState<ActiveEquation | null>(null);
  const [ready, setReady] = useState(false);
  const [status, setStatus] = useState('Loading equation metadata…');
  onErrorRef.current = onError;
  equationDocumentRef.current = equationDocument;
  readyRef.current = ready;

  useEffect(() => {
    let cancelled = false;
    setDocumentSource(initialDocument);
    setReady(false);
    setStatus('Loading equation metadata…');
    void loadEquationDocument(initialDocument).then((loaded) => {
      if (cancelled) return;
      setEquationDocument(loaded);
      setStatus(`${loaded.equations.length} equations ready to edit`);
    }).catch((error) => {
      if (cancelled) return;
      setStatus('Equation editing is unavailable');
      onErrorRef.current?.(error);
    });
    return () => { cancelled = true; };
  }, [initialDocument]);

  const closeEditor = useCallback(() => {
    activeMathRef.current = null;
    setActiveEquation(null);
  }, []);

  const cancelEditor = useCallback(() => {
    const editing = activeEquation?.mode === 'edit' ? activeEquation : null;
    const target = editing?.selectionTarget ?? null;
    closeEditor();
    if (!editing || !target || target.end.kind !== 'text') return;
    const end = { ...target.end, offset: editing.reference.paintedEndOffset };
    const caret: SelectionTarget = { ...target, start: end, end };
    requestAnimationFrame(() => {
      const activeEditor = editorRef.current?.getInstance()?.activeEditor as
        | { host?: SelectionTargetHost }
        | null
        | undefined;
      const selectionTargets = activeEditor?.host?.getHandles().editing?.selectionTargets;
      const focusEditor = () => {
        editorHostRef.current?.querySelector<HTMLTextAreaElement>('textarea[data-v2-ime-host="true"]')
          ?.focus({ preventScroll: true });
      };
      const placement = selectionTargets?.applyAndWait?.(caret);
      if (placement) void placement.finally(focusEditor);
      else focusEditor();
    });
  }, [activeEquation, closeEditor, editorRef]);

  const getLiveDocumentApi = useCallback(() => (
    editorRef.current?.getInstance()?.activeEditor?.doc as BrowserDocumentApi | null | undefined
  ), [editorRef]);

  const refreshEquationMetadata = useCallback(async (doc: BrowserDocumentApi) => {
    const exported = await doc.export.toDocx({ mode: 'review-preserving' });
    const bytes = bytesFromBase64(exported.contentBase64);
    const loaded = await loadEquationDocument(new Blob([bytes], {
      type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    }));
    setEquationDocument(loaded);
  }, []);

  const restoreScroll = useCallback(() => {
    const snapshot = pendingScrollRef.current;
    const host = editorHostRef.current;
    if (snapshot && host) restoreScrollSnapshot(host, snapshot);
  }, []);

  const scheduleScrollRestore = useCallback(() => {
    if (!pendingScrollRef.current) return;
    requestAnimationFrame(() => requestAnimationFrame(restoreScroll));
    window.setTimeout(restoreScroll, 100);
    window.setTimeout(() => {
      restoreScroll();
      pendingScrollRef.current = null;
    }, 350);
  }, [restoreScroll]);

  useEffect(() => {
    const host = editorHostRef.current;
    if (!host || !ready || !equationDocument) return;
    const openEditor = async (event: MouseEvent) => {
      const target = event.target;
      if (!(target instanceof Element)) return;
      const math = target.closest('math');
      if (!math || !host.contains(math)) return;
      const match = equationReferenceFromClick(math, equationDocument);
      if (!match) return;
      const doc = getLiveDocumentApi();
      await new Promise<void>((resolve) => requestAnimationFrame(() => resolve()));
      const selection = doc ? await doc.selection.current({ includeText: false }) : null;
      const selectionTarget = equationRangeFromCaret(selection?.selectionTarget ?? null, match.reference)
        ?? equationRangeFromBlock(match.blockId, match.reference);
      activeMathRef.current = math;
      setActiveEquation({
        equationNumber: match.equationNumber,
        latex: match.reference.latex,
        mode: 'edit',
        position: getEquationEditorPosition(math.getBoundingClientRect()),
        reference: match.reference,
        selectionTarget,
      });
      setStatus(
        selectionTarget
          ? `Equation ${match.equationNumber} is ready for an in-place update`
          : `Equation ${match.equationNumber} position could not be resolved; DOCX fallback will be used`,
      );
    };
    host.addEventListener('click', openEditor);
    return () => host.removeEventListener('click', openEditor);
  }, [equationDocument, getLiveDocumentApi, ready]);

  const openInsertionEditor = useCallback((context: ToolbarSelectContext) => {
    const host = editorHostRef.current;
    if (!host || !readyRef.current || !equationDocumentRef.current) return;
    const selection = context.ui.selection.current();
    const target = selection?.selectionTarget;
    if (!selection?.empty || !target || target.start.kind !== 'text' || target.end.kind !== 'text') {
      setStatus('Place a collapsed cursor in the document before inserting an equation');
      return;
    }
    const resolved = insertionPointFromSelection(host, target.end.blockId, target.end.offset);
    const anchorRect = context.ui.selection.getAnchorRect({ placement: 'end' });
    if (!resolved || !anchorRect) {
      setStatus('The current cursor position cannot accept an equation');
      return;
    }
    const bounds = new DOMRect(anchorRect.left, anchorRect.top, Math.max(anchorRect.width, 1), Math.max(anchorRect.height, 1));
    activeMathRef.current = resolved.fragment;
    setActiveEquation({
      insertionPoint: resolved.insertionPoint,
      latex: '',
      mode: 'insert',
      position: getEquationEditorPosition(bounds),
      selectionTarget: target,
    });
  }, []);

  const toolbarConfig = useMemo<ToolbarConfig>(() => ({
    customItems: [
      {
        attributes: { ariaLabel: 'Insert equation' },
        icon: '<svg viewBox="0 0 24 24" aria-hidden="true"><text x="12" y="18" text-anchor="middle" font-size="21" font-family="serif">∑</text></svg>',
        id: 'insert-equation',
        onSelect: openInsertionEditor,
        region: 'center',
        tooltip: 'Insert equation',
        type: 'button',
      },
    ],
    overflow: 'visible',
  }), [openInsertionEditor]);

  const applyEquation = useCallback(async (latex: string) => {
    if (!activeEquation || !equationDocument) return;
    const isInsertion = activeEquation.mode === 'insert';
    const isDeletion = activeEquation.mode === 'edit' && !latex.trim();
    if (isInsertion && !latex.trim()) {
      closeEditor();
      return;
    }
    const equationNumber = activeEquation.mode === 'edit' ? activeEquation.equationNumber : null;
    setStatus(isInsertion ? 'Inserting equation…' : isDeletion ? `Deleting equation ${equationNumber}…` : `Updating equation ${equationNumber}…`);
    try {
      const doc = getLiveDocumentApi();
      const liveTarget = activeEquation.selectionTarget;
      const replacementOoxml = isDeletion
        ? null
        : activeEquation.mode === 'edit' && activeEquation.reference.isDisplay
          ? mathParagraphOoxmlFromLatex(latex)
          : equationOoxmlFromLatex(latex);
      if (doc && (activeEquation.mode === 'insert' || isRangeSelection(liveTarget))) {
        try {
          const receipt = activeEquation.mode === 'insert'
            ? await doc.invoke({
                operationId: 'insert',
                input: {
                  namespaces: OOXML_NAMESPACES,
                  target: liveTarget,
                  type: 'ooxml',
                  value: replacementOoxml!,
                },
              })
            : isDeletion
              ? await doc.invoke({ operationId: 'delete', input: { target: liveTarget } })
              : await doc.invoke({
                operationId: 'replace',
                input: {
                  namespaces: OOXML_NAMESPACES,
                  target: liveTarget,
                  type: 'ooxml',
                  value: replacementOoxml!,
                },
              });
          if (successfulReceipt(receipt)) {
            closeEditor();
            await refreshEquationMetadata(doc);
            setStatus(
              isInsertion
                ? 'Equation inserted in place with SuperDoc'
                : isDeletion
                  ? `Equation ${equationNumber} deleted in place with SuperDoc`
                  : `Equation ${equationNumber} updated in place with SuperDoc`,
            );
            return;
          }
          console.info('SuperDoc declined the native equation mutation; using the DOCX fallback.', JSON.stringify(receipt));
        } catch (error) {
          console.info('SuperDoc could not apply the native equation mutation; using the DOCX fallback.', error);
        }
      }

      const host = editorHostRef.current;
      if (host) pendingScrollRef.current = captureScrollSnapshot(host, activeMathRef.current);
      const updated = activeEquation.mode === 'insert'
        ? await insertEquation(equationDocument.bytes, activeEquation.insertionPoint, latex)
        : isDeletion
          ? await deleteEquation(equationDocument.bytes, activeEquation.reference)
          : await replaceEquation(equationDocument.bytes, activeEquation.reference, latex);
      setEquationDocument({ bytes: updated.bytes, equations: updated.equations });
      setDocumentSource(updated.file);
      setReady(false);
      closeEditor();
      setDocumentVersion((version) => version + 1);
      setStatus(
        isInsertion
          ? 'Equation inserted; reloading document…'
          : isDeletion
            ? `Equation ${equationNumber} deleted; reloading document…`
            : `Equation ${equationNumber} updated; reloading document…`,
      );
    } catch (error) {
      setStatus(error instanceof Error ? error.message : 'The equation could not be updated.');
      onErrorRef.current?.(error);
    }
  }, [activeEquation, closeEditor, equationDocument, getLiveDocumentApi, refreshEquationMetadata]);

  const handleReady = useCallback(() => {
    setReady(true);
    setStatus(`${equationDocument?.equations.length ?? 0} equations ready to edit`);
    scheduleScrollRestore();
  }, [equationDocument, scheduleScrollRestore]);

  return {
    documentSource,
    documentVersion,
    editor: activeEquation ? <EquationEditor equationNumber={activeEquation.mode === 'edit' ? activeEquation.equationNumber : undefined} initialLatex={activeEquation.latex} mode={activeEquation.mode} onCancel={cancelEditor} onDone={applyEquation} position={activeEquation.position} /> : null,
    editorHostClassName: 'equation-editing-host',
    editorHostRef,
    onReady: handleReady,
    ready,
    status,
    toolbarConfig,
  };
}
