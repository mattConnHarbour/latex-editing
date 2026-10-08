import { useMemo, useRef, useState } from 'react';
import { SuperDocEditor, type SuperDocRef } from '@superdoc/react';
import '@superdoc/react/style.css';
import { useEquationEditing } from './equation-handling';

const INITIAL_DOCUMENT = '/calculus-equations.docx';

function reportDocumentError({ error }: { error: unknown }) {
  console.error('SuperDoc could not open the equations fixture.', error);
}

export default function App() {
  const editorRef = useRef<SuperDocRef>(null);
  const exportingRef = useRef(false);
  const [exporting, setExporting] = useState(false);
  const equationHandler = useEquationEditing({
    initialDocument: INITIAL_DOCUMENT,
    onError: (error) => console.error('Equation editing failed.', error),
  });
  const editorUi = useMemo(
    () => ({ toolbar: equationHandler.toolbarConfig }),
    [equationHandler.toolbarConfig],
  );

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

  return (
    <main>
      <header className='demo-toolbar'>
        <div>
          <strong>Equation rendering demo</strong>
          <span>{equationHandler.status}</span>
        </div>
        <button disabled={!equationHandler.ready || exporting} onClick={() => void exportDocument()} type='button'>
          {exporting ? 'Exporting…' : 'Export DOCX'}
        </button>
      </header>
      <div className={`editor-host ${equationHandler.editorHostClassName}`} ref={equationHandler.editorHostRef}>
        <SuperDocEditor
          document={equationHandler.documentSource}
          key={equationHandler.documentVersion}
          onContentError={reportDocumentError}
          onException={reportDocumentError}
          onReady={equationHandler.onReady}
          ref={editorRef}
          ui={editorUi}
        />
      </div>
      {equationHandler.editor}
    </main>
  );
}
