import { useCallback, useEffect, useRef, useState, type ReactNode, type RefObject } from 'react';
import EquationEditor from './EquationEditor';
import { loadEquationDocument, replaceEquation, type EquationDocument, type EquationReference } from './docx-equation-transformer';
import { captureScrollSnapshot, getEquationEditorPosition, restoreScrollSnapshot, type ScrollSnapshot } from './equation-positioning';
import 'katex/dist/katex.min.css';
import './equation-editor.css';

type ActiveEquation = {
  equationNumber: number;
  latex: string;
  position: { left: number; top: number };
  reference: EquationReference;
};

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
};

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

export function useEquationEditing({ initialDocument, onError }: UseEquationEditingOptions): EquationEditingController {
  const editorHostRef = useRef<HTMLDivElement>(null);
  const activeMathRef = useRef<Element | null>(null);
  const pendingScrollRef = useRef<ScrollSnapshot | null>(null);
  const onErrorRef = useRef(onError);
  const [documentSource, setDocumentSource] = useState<string | Blob>(initialDocument);
  const [documentVersion, setDocumentVersion] = useState(0);
  const [equationDocument, setEquationDocument] = useState<EquationDocument | null>(null);
  const [activeEquation, setActiveEquation] = useState<ActiveEquation | null>(null);
  const [ready, setReady] = useState(false);
  const [status, setStatus] = useState('Loading equation metadata…');
  onErrorRef.current = onError;

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
        position: getEquationEditorPosition(math.getBoundingClientRect()),
        reference: match.reference,
      });
    };
    host.addEventListener('click', openEditor, true);
    return () => host.removeEventListener('click', openEditor, true);
  }, [equationDocument, ready]);

  const applyEquation = useCallback(async (latex: string) => {
    if (!activeEquation || !equationDocument) return;
    const equationNumber = activeEquation.equationNumber;
    setStatus(`Updating equation ${equationNumber}…`);
    try {
      const host = editorHostRef.current;
      if (host) pendingScrollRef.current = captureScrollSnapshot(host, activeMathRef.current);
      const updated = await replaceEquation(equationDocument.bytes, activeEquation.reference, latex);
      setEquationDocument({ bytes: updated.bytes, equations: updated.equations });
      setDocumentSource(updated.file);
      setReady(false);
      closeEditor();
      setDocumentVersion((version) => version + 1);
      setStatus(`Equation ${equationNumber} updated; reloading document…`);
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
    editor: activeEquation ? <EquationEditor equationNumber={activeEquation.equationNumber} initialLatex={activeEquation.latex} onCancel={closeEditor} onDone={applyEquation} position={activeEquation.position} /> : null,
    editorHostClassName: 'equation-editing-host',
    editorHostRef,
    onReady: handleReady,
    ready,
    status,
  };
}
