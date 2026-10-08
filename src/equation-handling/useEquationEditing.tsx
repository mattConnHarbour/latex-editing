import { useCallback, useEffect, useMemo, useRef, useState, type ReactNode, type RefObject } from 'react';
import type { ToolbarConfig } from '@superdoc/react';
import EquationEditor from './EquationEditor';
import {
  deleteEquation,
  insertEquation,
  loadEquationDocument,
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
  | { mode: 'edit'; equationNumber: number; reference: EquationReference }
  | { mode: 'insert'; insertionPoint: EquationInsertionPoint }
);

export type UseEquationEditingOptions = {
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
    return reference ? { equationNumber: equationNumber + 1, reference } : null;
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

export function useEquationEditing({ initialDocument, onError }: UseEquationEditingOptions): EquationEditingController {
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
    const openEditor = (event: MouseEvent) => {
      const target = event.target;
      if (!(target instanceof Element)) return;
      const math = target.closest('math');
      if (!math || !host.contains(math)) return;
      const match = equationReferenceFromClick(math, equationDocument);
      if (!match) return;
      event.preventDefault();
      event.stopPropagation();
      activeMathRef.current = math;
      setActiveEquation({
        equationNumber: match.equationNumber,
        latex: match.reference.latex,
        mode: 'edit',
        position: getEquationEditorPosition(math.getBoundingClientRect()),
        reference: match.reference,
      });
    };
    host.addEventListener('click', openEditor, true);
    return () => host.removeEventListener('click', openEditor, true);
  }, [equationDocument, ready]);

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
  }, [activeEquation, closeEditor, equationDocument]);

  const handleReady = useCallback(() => {
    setReady(true);
    setStatus(`${equationDocument?.equations.length ?? 0} equations ready to edit`);
    scheduleScrollRestore();
  }, [equationDocument, scheduleScrollRestore]);

  return {
    documentSource,
    documentVersion,
    editor: activeEquation ? <EquationEditor equationNumber={activeEquation.mode === 'edit' ? activeEquation.equationNumber : undefined} initialLatex={activeEquation.latex} mode={activeEquation.mode} onCancel={closeEditor} onDone={applyEquation} position={activeEquation.position} /> : null,
    editorHostClassName: 'equation-editing-host',
    editorHostRef,
    onReady: handleReady,
    ready,
    status,
    toolbarConfig,
  };
}
