import { useEffect, useMemo, useRef, useState } from 'react';
import katex from 'katex';

export type EquationEditorProps = {
  equationNumber?: number;
  initialLatex: string;
  mode?: 'edit' | 'insert';
  position: { left: number; top: number };
  onCancel: () => void;
  onDone: (latex: string) => Promise<void>;
};

export default function EquationEditor({
  equationNumber,
  initialLatex,
  mode = 'edit',
  position,
  onCancel,
  onDone,
}: EquationEditorProps) {
  const [latex, setLatex] = useState(initialLatex);
  const [previewLatex, setPreviewLatex] = useState(initialLatex);
  const [saving, setSaving] = useState(false);
  const textareaRef = useRef<HTMLTextAreaElement>(null);

  useEffect(() => {
    textareaRef.current?.focus();
    textareaRef.current?.select();
  }, []);

  useEffect(() => {
    const timeout = window.setTimeout(() => setPreviewLatex(latex), 180);
    return () => window.clearTimeout(timeout);
  }, [latex]);

  useEffect(() => {
    const closeOnEscape = (event: KeyboardEvent) => {
      if (event.key === 'Escape' && !saving) onCancel();
    };
    window.addEventListener('keydown', closeOnEscape);
    return () => window.removeEventListener('keydown', closeOnEscape);
  }, [onCancel, saving]);

  const preview = useMemo(() => {
    if (!previewLatex.trim()) return { error: null, html: '' };
    try {
      return {
        error: null,
        html: katex.renderToString(previewLatex, {
          displayMode: true,
          throwOnError: true,
          strict: 'error',
          trust: false,
          maxExpand: 1_000,
          maxSize: 20,
        }),
      };
    } catch (error) {
      return {
        error: error instanceof Error ? error.message : 'This expression is not valid LaTeX.',
        html: '',
      };
    }
  }, [previewLatex]);

  async function submit() {
    if (preview.error || saving) return;
    setSaving(true);
    try {
      await onDone(latex);
    } finally {
      setSaving(false);
    }
  }

  return (
    <div
      aria-labelledby='equation-editor-title'
      aria-modal='true'
      className='equation-editor'
      data-testid='equation-editor'
      role='dialog'
      style={{ left: position.left, top: position.top }}
    >
      <div className='equation-editor__heading'>
        <div>
          <span>{mode === 'insert' ? 'New equation' : `Equation ${equationNumber}`}</span>
          <h2 id='equation-editor-title'>{mode === 'insert' ? 'Insert equation' : 'Edit LaTeX'}</h2>
        </div>
        <button aria-label='Close equation editor' disabled={saving} onClick={onCancel} type='button'>
          ×
        </button>
      </div>

      <label htmlFor='equation-latex'>LaTeX expression</label>
      <textarea
        id='equation-latex'
        onChange={(event) => setLatex(event.target.value)}
        ref={textareaRef}
        rows={4}
        spellCheck={false}
        value={latex}
      />

      <div className='equation-editor__preview' data-testid='equation-preview'>
        <span>Preview</span>
        {preview.error ? (
          <p className='equation-editor__error' role='alert'>{preview.error}</p>
        ) : (
          <div dangerouslySetInnerHTML={{ __html: preview.html }} />
        )}
      </div>

      <div className='equation-editor__actions'>
        <button disabled={saving} onClick={onCancel} type='button'>Cancel</button>
        <button disabled={Boolean(preview.error) || saving} onClick={() => void submit()} type='button'>
          {saving ? 'Updating…' : mode === 'edit' && !latex.trim() ? 'Delete' : 'Done'}
        </button>
      </div>
    </div>
  );
}
