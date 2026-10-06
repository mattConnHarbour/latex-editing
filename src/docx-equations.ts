import { ommlToLatex } from 'dwml-ts';
import JSZip from 'jszip';
import katex from 'katex';
import { mml2omml } from 'mathml2omml-plus';

const DOCUMENT_XML_PATH = 'word/document.xml';
const MATH_NAMESPACE = 'http://schemas.openxmlformats.org/officeDocument/2006/math';
const WORD_NAMESPACE = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

export type EquationDocument = {
  bytes: Uint8Array;
  equations: EquationReference[];
};

export type EquationReference = {
  equationIndexInParagraph: number;
  latex: string;
  paragraphOrdinal: number;
};

function parseXml(xml: string): XMLDocument {
  const document = new DOMParser().parseFromString(xml, 'application/xml');
  const parserError = document.querySelector('parsererror');
  if (parserError) throw new Error(`The equation XML is invalid: ${parserError.textContent ?? 'unknown error'}`);
  return document;
}

function mathmlFromLatex(latex: string): string {
  const rendered = katex.renderToString(latex, {
    displayMode: true,
    output: 'mathml',
    throwOnError: true,
    strict: 'error',
    trust: false,
    maxExpand: 1_000,
    maxSize: 20,
  });
  const renderedDocument = new DOMParser().parseFromString(rendered, 'text/html');
  const math = renderedDocument.querySelector('math');
  if (!math) throw new Error('KaTeX did not produce MathML for this expression.');
  return new XMLSerializer().serializeToString(math);
}

async function readDocumentXml(bytes: Uint8Array): Promise<{ archive: JSZip; xml: string }> {
  const archive = await JSZip.loadAsync(bytes);
  const documentPart = archive.file(DOCUMENT_XML_PATH);
  if (!documentPart) throw new Error('This DOCX does not contain word/document.xml.');
  return { archive, xml: await documentPart.async('string') };
}

function inspectEquations(xml: string): EquationReference[] {
  const document = parseXml(xml);
  const body = document.getElementsByTagNameNS(WORD_NAMESPACE, 'body')[0];
  if (!body) throw new Error('This DOCX does not contain a Word document body.');
  const paragraphs = Array.from(body.children).filter(
    (element) => element.namespaceURI === WORD_NAMESPACE && element.localName === 'p',
  );
  const serializer = new XMLSerializer();
  return paragraphs.flatMap((paragraph, paragraphOrdinal) =>
    Array.from(paragraph.getElementsByTagNameNS(MATH_NAMESPACE, 'oMath')).map(
      (equation, equationIndexInParagraph) => ({
        equationIndexInParagraph,
        latex: ommlToLatex(serializer.serializeToString(equation)),
        paragraphOrdinal,
      }),
    ),
  );
}

export async function loadEquationDocument(source: string | Blob): Promise<EquationDocument> {
  const response = typeof source === 'string' ? await fetch(source) : null;
  if (response && !response.ok) throw new Error(`Could not load the DOCX (${response.status}).`);
  const buffer = response ? await response.arrayBuffer() : await (source as Blob).arrayBuffer();
  const bytes = new Uint8Array(buffer);
  const { xml } = await readDocumentXml(bytes);
  return { bytes, equations: inspectEquations(xml) };
}

export async function replaceEquation(
  bytes: Uint8Array,
  reference: Pick<EquationReference, 'equationIndexInParagraph' | 'paragraphOrdinal'>,
  latex: string,
): Promise<EquationDocument & { file: File }> {
  const mathml = mathmlFromLatex(latex);
  const replacementOmml = mml2omml(mathml);
  const { archive, xml } = await readDocumentXml(bytes);
  const document = parseXml(xml);
  const body = document.getElementsByTagNameNS(WORD_NAMESPACE, 'body')[0];
  const paragraphs = body
    ? Array.from(body.children).filter(
        (element) => element.namespaceURI === WORD_NAMESPACE && element.localName === 'p',
      )
    : [];
  const paragraph = paragraphs[reference.paragraphOrdinal];
  const target = paragraph?.getElementsByTagNameNS(MATH_NAMESPACE, 'oMath')[reference.equationIndexInParagraph];
  if (!target) throw new Error('The selected equation no longer exists in the document.');

  const replacementDocument = parseXml(replacementOmml);
  const replacement = replacementDocument.getElementsByTagNameNS(MATH_NAMESPACE, 'oMath')[0];
  if (!replacement) throw new Error('The converted equation did not contain an OMML math object.');

  target.replaceWith(document.importNode(replacement, true));
  archive.file(DOCUMENT_XML_PATH, new XMLSerializer().serializeToString(document));
  const nextBytes = await archive.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
  const fileBuffer = nextBytes.buffer.slice(nextBytes.byteOffset, nextBytes.byteOffset + nextBytes.byteLength) as ArrayBuffer;
  const file = new File([fileBuffer], 'calculus-equations-edited.docx', { type: DOCX_MIME });
  const equations = inspectEquations(await archive.file(DOCUMENT_XML_PATH)!.async('string'));
  return { bytes: nextBytes, equations, file };
}
