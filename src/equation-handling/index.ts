export { default as EquationEditor, type EquationEditorProps } from './EquationEditor';
export { deleteEquation, insertEquation, loadEquationDocument, replaceEquation, type EquationDocument, type EquationInsertionPoint, type EquationReference } from './docx-equation-transformer';
export { captureScrollSnapshot, getEquationEditorPosition, restoreScrollSnapshot, type EquationEditorPosition, type ScrollSnapshot } from './equation-positioning';
export { useEquationEditing, type EquationEditingController, type UseEquationEditingOptions } from './useEquationEditing';
