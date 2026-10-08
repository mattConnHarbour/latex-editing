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

export type EquationInsertionPoint = {
  offset: number;
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

function ommlFromLatex(latex: string): Element {
  const replacementDocument = parseXml(mml2omml(mathmlFromLatex(latex)));
  const equation = replacementDocument.getElementsByTagNameNS(MATH_NAMESPACE, 'oMath')[0];
  if (!equation) throw new Error('The converted equation did not contain an OMML math object.');
  return equation;
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

function documentParagraphs(document: XMLDocument): Element[] {
  const body = document.getElementsByTagNameNS(WORD_NAMESPACE, 'body')[0];
  return body
    ? Array.from(body.children).filter(
        (element) => element.namespaceURI === WORD_NAMESPACE && element.localName === 'p',
      )
    : [];
}

function setRunText(run: Element, text: string) {
  const textNodes = Array.from(run.getElementsByTagNameNS(WORD_NAMESPACE, 't'));
  for (const node of textNodes) node.remove();
  if (!text) return;
  const textNode = run.ownerDocument.createElementNS(WORD_NAMESPACE, 'w:t');
  if (/^\s|\s$/.test(text)) textNode.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
  textNode.textContent = text;
  run.append(textNode);
}

function topLevelParagraphChild(paragraph: Element, node: Node): Node | null {
  let current: Node | null = node;
  while (current?.parentNode && current.parentNode !== paragraph) current = current.parentNode;
  return current?.parentNode === paragraph ? current : null;
}

function paragraphSplitBoundary(paragraph: Element, offset: number): Node | null {
  const runs = Array.from(paragraph.getElementsByTagNameNS(WORD_NAMESPACE, 'r'));
  let consumed = 0;

  for (const run of runs) {
    const text = Array.from(run.getElementsByTagNameNS(WORD_NAMESPACE, 't'))
      .map((node) => node.textContent ?? '')
      .join('');
    const nextOffset = consumed + text.length;
    if (offset <= nextOffset) {
      const localOffset = Math.max(0, offset - consumed);
      const parent = run.parentNode;
      const topLevelChild = topLevelParagraphChild(paragraph, run);
      if (!parent || !topLevelChild) return null;
      if (localOffset === 0) return topLevelChild;
      if (localOffset === text.length) return topLevelChild.nextSibling;
      if (topLevelChild === run) {
        const trailingRun = run.cloneNode(true) as Element;
        setRunText(run, text.slice(0, localOffset));
        setRunText(trailingRun, text.slice(localOffset));
        parent.insertBefore(trailingRun, run.nextSibling);
        return trailingRun;
      }
      throw new Error('Equations cannot yet be inserted in the middle of nested text content.');
    }
    consumed = nextOffset;
  }

  return null;
}

function createDisplayEquationParagraph(document: XMLDocument, equation: Element): Element {
  const paragraph = document.createElementNS(WORD_NAMESPACE, 'w:p');
  const paragraphProperties = document.createElementNS(WORD_NAMESPACE, 'w:pPr');
  const spacing = document.createElementNS(WORD_NAMESPACE, 'w:spacing');
  spacing.setAttributeNS(WORD_NAMESPACE, 'w:before', '0');
  spacing.setAttributeNS(WORD_NAMESPACE, 'w:after', '140');
  paragraphProperties.append(spacing);
  paragraph.append(paragraphProperties);

  const mathParagraph = document.createElementNS(MATH_NAMESPACE, 'm:oMathPara');
  const mathParagraphProperties = document.createElementNS(MATH_NAMESPACE, 'm:oMathParaPr');
  const justification = document.createElementNS(MATH_NAMESPACE, 'm:jc');
  justification.setAttributeNS(MATH_NAMESPACE, 'm:val', 'centerGroup');
  mathParagraphProperties.append(justification);
  mathParagraph.append(mathParagraphProperties, document.importNode(equation, true));
  paragraph.append(mathParagraph);
  return paragraph;
}

function insertDisplayEquationAtTextOffset(paragraph: Element, offset: number, equation: Element) {
  const document = paragraph.ownerDocument;
  const parent = paragraph.parentNode;
  if (!parent) throw new Error('The selected paragraph is detached from the document.');
  const boundary = paragraphSplitBoundary(paragraph, offset);
  const equationParagraph = createDisplayEquationParagraph(document, equation);

  if (!boundary) {
    parent.insertBefore(equationParagraph, paragraph.nextSibling);
    return;
  }

  const firstContent = Array.from(paragraph.childNodes).find(
    (node) => !(node instanceof Element && node.namespaceURI === WORD_NAMESPACE && node.localName === 'pPr'),
  );
  if (boundary === firstContent) {
    parent.insertBefore(equationParagraph, paragraph);
    return;
  }

  const trailingParagraph = document.createElementNS(WORD_NAMESPACE, 'w:p');
  const paragraphProperties = Array.from(paragraph.children).find(
    (element) => element.namespaceURI === WORD_NAMESPACE && element.localName === 'pPr',
  );
  if (paragraphProperties) trailingParagraph.append(paragraphProperties.cloneNode(true));
  let trailingNode: Node | null = boundary;
  while (trailingNode) {
    const nextNode: Node | null = trailingNode.nextSibling;
    trailingParagraph.append(trailingNode);
    trailingNode = nextNode;
  }

  parent.insertBefore(equationParagraph, paragraph.nextSibling);
  parent.insertBefore(trailingParagraph, equationParagraph.nextSibling);
}

async function createUpdatedDocument(archive: JSZip, document: XMLDocument) {
  archive.file(DOCUMENT_XML_PATH, new XMLSerializer().serializeToString(document));
  const nextBytes = await archive.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
  const fileBuffer = nextBytes.buffer.slice(nextBytes.byteOffset, nextBytes.byteOffset + nextBytes.byteLength) as ArrayBuffer;
  const file = new File([fileBuffer], 'calculus-equations-edited.docx', { type: DOCX_MIME });
  const equations = inspectEquations(await archive.file(DOCUMENT_XML_PATH)!.async('string'));
  return { bytes: nextBytes, equations, file };
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
  const { archive, xml } = await readDocumentXml(bytes);
  const document = parseXml(xml);
  const paragraphs = documentParagraphs(document);
  const paragraph = paragraphs[reference.paragraphOrdinal];
  const target = paragraph?.getElementsByTagNameNS(MATH_NAMESPACE, 'oMath')[reference.equationIndexInParagraph];
  if (!target) throw new Error('The selected equation no longer exists in the document.');

  target.replaceWith(document.importNode(ommlFromLatex(latex), true));
  return createUpdatedDocument(archive, document);
}

export async function deleteEquation(
  bytes: Uint8Array,
  reference: Pick<EquationReference, 'equationIndexInParagraph' | 'paragraphOrdinal'>,
): Promise<EquationDocument & { file: File }> {
  const { archive, xml } = await readDocumentXml(bytes);
  const document = parseXml(xml);
  const paragraph = documentParagraphs(document)[reference.paragraphOrdinal];
  const target = paragraph?.getElementsByTagNameNS(MATH_NAMESPACE, 'oMath')[reference.equationIndexInParagraph];
  if (!target) throw new Error('The selected equation no longer exists in the document.');

  const mathParagraph = target.parentElement;
  if (
    mathParagraph?.namespaceURI === MATH_NAMESPACE &&
    mathParagraph.localName === 'oMathPara' &&
    mathParagraph.getElementsByTagNameNS(MATH_NAMESPACE, 'oMath').length === 1
  ) {
    mathParagraph.remove();
  } else {
    target.remove();
  }
  return createUpdatedDocument(archive, document);
}

export async function insertEquation(
  bytes: Uint8Array,
  insertionPoint: EquationInsertionPoint,
  latex: string,
): Promise<EquationDocument & { file: File }> {
  const { archive, xml } = await readDocumentXml(bytes);
  const document = parseXml(xml);
  const paragraph = documentParagraphs(document)[insertionPoint.paragraphOrdinal];
  if (!paragraph) throw new Error('The selected paragraph no longer exists in the document.');

  insertDisplayEquationAtTextOffset(paragraph, insertionPoint.offset, ommlFromLatex(latex));
  return createUpdatedDocument(archive, document);
}
