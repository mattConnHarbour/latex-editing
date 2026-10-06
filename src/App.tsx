import { useCallback, useEffect, useRef, useState } from 'react';
import { SuperDocEditor, type SuperDocRef } from '@superdoc/react';
import '@superdoc/react/style.css';
import 'katex/dist/katex.min.css';
import EquationEditor from './EquationEditor';
import {
  loadEquationDocument,
  replaceEquation,
  type EquationDocument,
  type EquationReference,
} from './docx-equations';

const INITIAL_DOCUMENT = '/calculus-equations.docx';
const EDITOR_WIDTH = 440;
const EDITOR_ESTIMATED_HEIGHT = 470;

type ActiveEquation = {
  equationNumber: number;
  latex: string;
  position: { left: number; top: number };
  reference: EquationReference;
};

type ScrollSnapshot = {
  containers: Array<{ path: number[]; left: number; top: number }>;
  windowX: number;
  windowY: number;
};

function reportDocumentError({ error }: { error: unknown }) {
  console.error('SuperDoc could not open the equations fixture.', error);
}

function clamp(value: number, minimum: number, maximum: number) {
  return Math.min(Math.max(value, minimum), Math.max(minimum, maximum));
}

function pathFromHost(host: HTMLElement, element: HTMLElement): number[] | null {
  const path: number[] = [];
  let current: HTMLElement | null = element;
  while (current && current !== host) {
    const parent: HTMLElement | null = current.parentElement;
    if (!parent) return null;
    path.unshift(Array.from(parent.children).indexOf(current));
    current = parent;
  }
  return current === host ? path : null;
}

function elementFromPath(host: HTMLElement, path: number[]): HTMLElement | null {
  let current: Element = host;
  for (const index of path) {
    const next: Element | undefined = current.children[index];
    if (!(next instanceof HTMLElement)) return null;
    current = next;
  }
  return current instanceof HTMLElement ? current : null;
}

export default function App() {
  const editorRef = useRef<SuperDocRef>(null);
  const editorHostRef = useRef<HTMLDivElement>(null);
  const activeMathRef = useRef<Element | null>(null);
  const pendingScrollRef = useRef<ScrollSnapshot | null>(null);
  const exportingRef = useRef(false);
  const [documentSource, setDocumentSource] = useState<string | File>(INITIAL_DOCUMENT);
  const [documentVersion, setDocumentVersion] = useState(0);
  const [equationDocument, setEquationDocument] = useState<EquationDocument | null>(null);
  const [activeEquation, setActiveEquation] = useState<ActiveEquation | null>(null);
  const [ready, setReady] = useState(false);
  const [exporting, setExporting] = useState(false);
  const [status, setStatus] = useState('Loading equation metadata…');

  useEffect(() => {
    let cancelled = false;
    void loadEquationDocument(INITIAL_DOCUMENT)
      .then((loaded) => {
        if (cancelled) return;
        setEquationDocument(loaded);
        setStatus(`${loaded.equations.length} equations ready to edit`);
      })
      .catch((error) => {
        if (cancelled) return;
        console.error('Could not inspect the equation fixture.', error);
        setStatus('Equation editing is unavailable');
      });
    return () => {
      cancelled = true;
    };
  }, []);

  const closeEquationEditor = useCallback(() => {
    activeMathRef.current = null;
    setActiveEquation(null);
  }, []);

  function captureScrollPosition(): ScrollSnapshot {
    const host = editorHostRef.current;
    const containers: ScrollSnapshot['containers'] = [];
    let current = activeMathRef.current?.parentElement ?? null;
    while (host && current) {
      if (current.scrollTop !== 0 || current.scrollLeft !== 0) {
        const path = pathFromHost(host, current);
        if (path) containers.push({ path, left: current.scrollLeft, top: current.scrollTop });
      }
      if (current === host) break;
      current = current.parentElement;
    }
    return { containers, windowX: window.scrollX, windowY: window.scrollY };
  }

  function restoreScrollPosition() {
    const snapshot = pendingScrollRef.current;
    const host = editorHostRef.current;
    if (!snapshot || !host) return;
    window.scrollTo(snapshot.windowX, snapshot.windowY);
    for (const container of snapshot.containers) {
      elementFromPath(host, container.path)?.scrollTo(container.left, container.top);
    }
  }

  function scheduleScrollRestore() {
    if (!pendingScrollRef.current) return;
    const restore = () => restoreScrollPosition();
    requestAnimationFrame(() => requestAnimationFrame(restore));
    window.setTimeout(restore, 100);
    window.setTimeout(() => {
      restore();
      pendingScrollRef.current = null;
    }, 350);
  }

  useEffect(() => {
    const host = editorHostRef.current;
    if (!host || !ready || !equationDocument) return;

    const openEquationEditor = (event: MouseEvent) => {
      const target = event.target;
      if (!(target instanceof Element)) return;
      const math = target.closest('math');
      if (!math || !host.contains(math)) return;

      const fragment = math.closest<HTMLElement>('.superdoc-fragment[data-source-anchor]');
      if (!fragment) return;
      let paragraphOrdinal: number | null = null;
      try {
        const anchor = JSON.parse(fragment.dataset.sourceAnchor ?? '{}') as {
          sourceRef?: { xpathLikePath?: string };
        };
        const match = /body\/w:p\[ordinal=(\d+)\]/.exec(anchor.sourceRef?.xpathLikePath ?? '');
        if (match) paragraphOrdinal = Number(match[1]);
      } catch {
        return;
      }
      if (paragraphOrdinal === null) return;
      const equationIndexInParagraph = Array.from(fragment.querySelectorAll('math')).indexOf(math);
      const equationNumber = equationDocument.equations.findIndex(
        (equation) =>
          equation.paragraphOrdinal === paragraphOrdinal &&
          equation.equationIndexInParagraph === equationIndexInParagraph,
      );
      const reference = equationDocument.equations[equationNumber];
      if (!reference) return;

      event.preventDefault();
      event.stopPropagation();
      const bounds = math.getBoundingClientRect();
      const preferredLeft = bounds.right + 16;
      const fallbackLeft = bounds.left - EDITOR_WIDTH - 16;
      const left = clamp(
        preferredLeft + EDITOR_WIDTH <= window.innerWidth - 12 ? preferredLeft : fallbackLeft,
        12,
        window.innerWidth - EDITOR_WIDTH - 12,
      );
      const top = clamp(bounds.top, 72, window.innerHeight - EDITOR_ESTIMATED_HEIGHT - 12);
      activeMathRef.current = math;
      setActiveEquation({
        equationNumber: equationNumber + 1,
        latex: reference.latex,
        position: { left, top },
        reference,
      });
    };

    host.addEventListener('click', openEquationEditor, true);
    return () => host.removeEventListener('click', openEquationEditor, true);
  }, [equationDocument, ready]);

  async function exportDocument() {
    if (exportingRef.current) return;
    exportingRef.current = true;
    setExporting(true);
    try {
      await editorRef.current?.getInstance()?.export({
        exportType: ['docx'],
        exportedName: 'calculus-equations-edited',
      });
    } catch (error) {
      console.error('SuperDoc could not export the equations fixture.', error);
    } finally {
      exportingRef.current = false;
      setExporting(false);
    }
  }

  async function applyEquation(latex: string) {
    if (!activeEquation || !equationDocument) return;
    const { equationNumber } = activeEquation;
    setStatus(`Updating equation ${equationNumber}…`);
    try {
      pendingScrollRef.current = captureScrollPosition();
      const updated = await replaceEquation(equationDocument.bytes, activeEquation.reference, latex);
      setEquationDocument({ bytes: updated.bytes, equations: updated.equations });
      setDocumentSource(updated.file);
      setReady(false);
      activeMathRef.current = null;
      setActiveEquation(null);
      setDocumentVersion((version) => version + 1);
      setStatus(`Equation ${equationNumber} updated; reloading document…`);
    } catch (error) {
      setStatus(error instanceof Error ? error.message : 'The equation could not be updated.');
    }
  }

  return (
    <main>
      <header className='demo-toolbar'>
        <div>
          <strong>Equation rendering demo</strong>
          <span>{status}</span>
        </div>
        <button disabled={!ready || exporting} onClick={() => void exportDocument()} type='button'>
          {exporting ? 'Exporting…' : 'Export DOCX'}
        </button>
      </header>
      <div className='editor-host' ref={editorHostRef}>
        <SuperDocEditor
          document={documentSource}
          key={documentVersion}
          onContentError={reportDocumentError}
          onException={reportDocumentError}
          onReady={() => {
            setReady(true);
            setStatus(`${equationDocument?.equations.length ?? 0} equations ready to edit`);
            scheduleScrollRestore();
          }}
          ref={editorRef}
        />
      </div>
      {activeEquation ? (
        <EquationEditor
          equationNumber={activeEquation.equationNumber}
          initialLatex={activeEquation.latex}
          onCancel={closeEquationEditor}
          onDone={applyEquation}
          position={activeEquation.position}
        />
      ) : null}
    </main>
  );
}
