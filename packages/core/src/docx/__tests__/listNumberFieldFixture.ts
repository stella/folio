/**
 * A package whose numbered paragraphs carry inline `LISTNUM` fields, and a
 * reading of saved paragraph markup that does not go through the parser.
 */

import { escapeXmlAttribute, escapeXmlText } from "@stll/docx-core";
import JSZip from "jszip";
import type { Node as PMNode } from "prosemirror-model";
import type { EditorState } from "prosemirror-state";

import { toFlowBlocks } from "../../layout-bridge/convert/toFlowBlocks";
import { toProseDoc } from "../../prosemirror/conversion/toProseDoc";
import type { ComplexField, Document, Paragraph } from "../../types/document";
import { foldedListNumberOf, isFoldedListNumber } from "../foldedListNumberFields";
import { parseDocx } from "../parser";
import { repackDocx } from "../rezip";
import { attemptSelectiveSave } from "../selectiveSave";

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

const CONTENT_TYPES = `${XML_DECLARATION}
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
  <Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
</Types>`;

const PACKAGE_RELS = `${XML_DECLARATION}
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
</Relationships>`;

const DOCUMENT_RELS = `${XML_DECLARATION}
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>
</Relationships>`;

const CORE_PROPERTIES = `${XML_DECLARATION}
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>List number fields</dc:title>
  <dcterms:modified xsi:type="dcterms:W3CDTF">2024-01-01T00:00:00.000Z</dcterms:modified>
</cp:coreProperties>`;

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

/** How the host paragraph's own level draws its marker. */
export type MarkerKind = "decimal" | "percent" | "symbol";

export const MARKER_KINDS: readonly MarkerKind[] = ["decimal", "percent", "symbol"];

/** The level-1 marker each kind declares, before any field is folded in. */
const LEVEL_MARKER: Record<MarkerKind, string> = {
  decimal: "%1.%2",
  // `%%` is a literal percent sign that no level placeholder owns.
  percent: "%1.%2 %%",
  symbol: "\uF0B7",
};

const levelOneXml = (kind: MarkerKind): string => {
  const format = kind === "symbol" ? "bullet" : "decimal";
  const fonts =
    kind === "symbol"
      ? '<w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/></w:rPr>'
      : "";
  return (
    `<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="${format}"/>` +
    `<w:lvlText w:val="${escapeXmlAttribute(LEVEL_MARKER[kind])}"/><w:lvlJc w:val="left"/>` +
    `<w:pPr><w:ind w:left="720" w:hanging="720"/></w:pPr>${fonts}</w:lvl>`
  );
};

/** One numbering instance per marker kind; `numId` is the kind's index plus one. */
const numberingXml = (): string => {
  const abstracts = MARKER_KINDS.map(
    (kind, index) =>
      `<w:abstractNum w:abstractNumId="${index}"><w:multiLevelType w:val="multilevel"/>` +
      `<w:lvl w:ilvl="0"><w:start w:val="7"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="720"/></w:pPr></w:lvl>` +
      levelOneXml(kind) +
      `<w:lvl w:ilvl="2"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="(%3)"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="1440" w:hanging="720"/></w:pPr></w:lvl>` +
      `</w:abstractNum>`,
  ).join("");
  const instances = MARKER_KINDS.map(
    (_, index) => `<w:num w:numId="${index + 1}"><w:abstractNumId w:val="${index}"/></w:num>`,
  ).join("");
  return `${XML_DECLARATION}<w:numbering ${W}>${abstracts}${instances}</w:numbering>`;
};

const numIdOf = (kind: MarkerKind): number => MARKER_KINDS.indexOf(kind) + 1;

/** Zero-width markup that sits between a field and the tab after it. */
export type GapMarker = "bookmark" | "bookmarkStart" | "comment";

export type ResultFormatting = "plain" | "bold" | "symbol";

export type FieldSpec = {
  /** The instruction as authored, spacing included. */
  instruction: string;
  /** The cached display, empty for none; a `\t` in it is a `w:tab` inside the result run. */
  result: string;
  formatting: ResultFormatting;
  /** Text in the paragraph ahead of this field. */
  before: string;
  gap: GapMarker[];
  tab: boolean;
};

export type ParagraphSpec = {
  paraId: string;
  marker: MarkerKind;
  fields: FieldSpec[];
  /** Text after the last field. */
  body: string;
};

const RESULT_PROPERTIES: Record<ResultFormatting, string> = {
  plain: "",
  bold: "<w:rPr><w:b/></w:rPr>",
  symbol: '<w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol"/></w:rPr>',
};

const textRun = (text: string): string =>
  text === "" ? "" : `<w:r><w:t xml:space="preserve">${escapeXmlText(text)}</w:t></w:r>`;

const resultRun = (field: FieldSpec): string => {
  // A field that cached nothing has no result run at all.
  if (field.result === "") {
    return "";
  }
  const pieces = field.result
    .split("\t")
    .map((piece) => (piece === "" ? "" : `<w:t xml:space="preserve">${escapeXmlText(piece)}</w:t>`))
    .join("<w:tab/>");
  return `<w:r>${RESULT_PROPERTIES[field.formatting]}${pieces}</w:r>`;
};

type MarkupIds = { next: number; comments: number[] };

const fieldXml = (field: FieldSpec, ids: MarkupIds): { inline: string; trailing: string } => {
  let gap = "";
  let trailing = "";
  for (const marker of field.gap) {
    const id = ids.next++;
    switch (marker) {
      case "bookmark":
        gap += `<w:bookmarkStart w:id="${id}" w:name="mark${id}"/><w:bookmarkEnd w:id="${id}"/>`;
        break;
      case "bookmarkStart":
        gap += `<w:bookmarkStart w:id="${id}" w:name="mark${id}"/>`;
        trailing += `<w:bookmarkEnd w:id="${id}"/>`;
        break;
      case "comment":
        ids.comments.push(id);
        gap += `<w:commentRangeStart w:id="${id}"/>`;
        trailing += `<w:commentRangeEnd w:id="${id}"/><w:r><w:commentReference w:id="${id}"/></w:r>`;
        break;
      default: {
        const unhandled: never = marker;
        throw new Error(`Unhandled gap marker ${String(unhandled)}`);
      }
    }
  }
  const inline =
    textRun(field.before) +
    `<w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
    `<w:r><w:instrText xml:space="preserve">${escapeXmlText(field.instruction)}</w:instrText></w:r>` +
    `<w:r><w:fldChar w:fldCharType="separate"/></w:r>` +
    resultRun(field) +
    `<w:r><w:fldChar w:fldCharType="end"/></w:r>` +
    gap +
    (field.tab ? "<w:r><w:tab/></w:r>" : "");
  return { inline, trailing };
};

const paragraphXml = (spec: ParagraphSpec, ids: MarkupIds, authored?: string): string => {
  let inline = "";
  let trailing = "";
  for (const field of spec.fields) {
    const built = fieldXml(field, ids);
    inline += built.inline;
    trailing += built.trailing;
  }
  return (
    `<w:p w14:paraId="${spec.paraId}"><w:pPr><w:numPr><w:ilvl w:val="1"/>` +
    `<w:numId w:val="${numIdOf(spec.marker)}"/></w:numPr></w:pPr>` +
    `${authored ?? `${inline}${textRun(spec.body)}${trailing}`}</w:p>`
  );
};

export const PLAIN_PARAGRAPH_ID = "20000009";

type FixtureOptions = {
  /** By `paraId`, inline markup to use in place of what the paragraph's spec would write. */
  authored?: Readonly<Record<string, string>>;
  /** Put the plain paragraph ahead of the numbered ones rather than after them. */
  plainFirst?: boolean;
};

/** A package of the given numbered paragraphs and one plain paragraph, which closes it. */
export const listNumberFieldDocx = (
  paragraphs: readonly ParagraphSpec[],
  { authored = {}, plainFirst = false }: FixtureOptions = {},
): Promise<ArrayBuffer> => {
  const ids: MarkupIds = { next: 1, comments: [] };
  const numbered = paragraphs
    .map((spec) => paragraphXml(spec, ids, authored[spec.paraId]))
    .join("");
  const plain = `<w:p w14:paraId="${PLAIN_PARAGRAPH_ID}"><w:r><w:t>Plain.</w:t></w:r></w:p>`;
  const documentXml =
    `${XML_DECLARATION}<w:document ${W} xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
    `<w:body>${plainFirst ? plain + numbered : numbered + plain}` +
    `<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>`;
  const commentsXml =
    `${XML_DECLARATION}<w:comments ${W}>` +
    ids.comments
      .map(
        (id) =>
          `<w:comment w:id="${id}" w:author="Reviewer" w:date="2026-01-01T00:00:00Z" w:initials="R"><w:p><w:r><w:t>Note ${id}</w:t></w:r></w:p></w:comment>`,
      )
      .join("") +
    `</w:comments>`;

  const zip = new JSZip();
  zip.file("[Content_Types].xml", CONTENT_TYPES);
  zip.file("_rels/.rels", PACKAGE_RELS);
  zip.file("word/_rels/document.xml.rels", DOCUMENT_RELS);
  zip.file("word/document.xml", documentXml);
  zip.file("word/numbering.xml", numberingXml());
  zip.file("word/comments.xml", commentsXml);
  zip.file("docProps/core.xml", CORE_PROPERTIES);
  return zip.generateAsync({ type: "arraybuffer" });
};

export const documentXmlOf = async (buffer: ArrayBuffer): Promise<string> => {
  const text = await (await JSZip.loadAsync(buffer)).file("word/document.xml")?.async("text");
  if (text === undefined) {
    throw new Error("The package has no word/document.xml");
  }
  return text;
};

const INLINE_TOKEN =
  /<w:fldChar\b[^>]*w:fldCharType="(?<fieldChar>\w+)"|<w:instrText\b[^>]*>(?<code>[^<]*)<\/w:instrText>|<w:t(?:\s[^>]*)?>(?<text>[^<]*)<\/w:t>|<w:(?<empty>tab|bookmarkStart|bookmarkEnd|commentRangeStart|commentRangeEnd|commentReference)\b/gu;

/**
 * The inline sequence of one `w:p`, read straight off its markup: field
 * characters, instruction text, text, tabs and range markers in the order
 * they were written. Adjacent text is joined, since where a run boundary
 * falls inside plain text carries no meaning. Text is compared as written:
 * the fixture uses no character that markup escapes.
 */
export const inlineTokens = (paragraphMarkup: string): string[] => {
  const inline = paragraphMarkup.replace(/<w:pPr>.*?<\/w:pPr>/su, "");
  const tokens: string[] = [];
  for (const match of inline.matchAll(INLINE_TOKEN)) {
    const groups = match.groups ?? {};
    if (groups["fieldChar"] !== undefined) {
      tokens.push(`fldChar:${groups["fieldChar"]}`);
    } else if (groups["code"] !== undefined) {
      tokens.push(`code:${groups["code"]}`);
    } else if (groups["text"] !== undefined) {
      const text = groups["text"];
      const last = tokens.at(-1);
      if (last?.startsWith("text:")) {
        tokens[tokens.length - 1] = last + text;
      } else if (text !== "") {
        tokens.push(`text:${text}`);
      }
    } else if (groups["empty"] !== undefined) {
      tokens.push(groups["empty"]);
    }
  }
  return tokens;
};

/** The markup of the `w:p` carrying `paraId`. */
export const paragraphMarkupOf = (documentXml: string, paraId: string): string => {
  const paragraph = documentXml
    .match(/<w:p[ >].*?<\/w:p>/gsu)
    ?.find((markup) => markup.includes(`w14:paraId="${paraId}"`));
  if (paragraph === undefined) {
    throw new Error(`The document has no paragraph ${paraId}`);
  }
  return paragraph;
};

/** What {@link inlineTokens} reads off a paragraph authored from `spec`. */
export const expectedTokens = (spec: ParagraphSpec): string[] =>
  inlineTokens(paragraphXml(spec, { next: 1, comments: [] }));

/**
 * `tokens` with everything after the last text put in a fixed order: the
 * range markers that close at the paragraph's end have no order among them
 * that a save owes the file.
 */
export const withSettledTail = (tokens: readonly string[]): string[] => {
  const lastText = tokens.findLastIndex((token) => token.startsWith("text:"));
  return [...tokens.slice(0, lastText + 1), ...tokens.slice(lastText + 1).toSorted()];
};

/**
 * The fields of `spec` the reader folds into the marker: the ones that open
 * the paragraph, when its marker has text to show them after and at least one
 * of them cached a display.
 */
export const foldedFieldsOf = (spec: ParagraphSpec): FieldSpec[] => {
  if (spec.marker === "symbol") {
    return [];
  }
  const opening: FieldSpec[] = [];
  for (const field of spec.fields) {
    if (field.before !== "") {
      break;
    }
    opening.push(field);
  }
  return opening.some((field) => field.result.replaceAll("\t", "") !== "") ? opening : [];
};

/** The cached display the marker shows for `fields`, joined as the marker joins it. */
export const cachedDisplay = (fields: readonly FieldSpec[]): string =>
  fields
    .map((field) => field.result.replaceAll("\t", ""))
    .filter((text) => text !== "")
    .join(" ");

/** The cached result of every field in a paragraph's markup, in order, as one line. */
export const fieldResultsInFile = (tokens: readonly string[]): string => {
  const results: string[] = [];
  let inResult = false;
  for (const token of tokens) {
    if (token === "fldChar:separate") {
      inResult = true;
      results.push("");
    } else if (token === "fldChar:end") {
      inResult = false;
    } else if (inResult && token.startsWith("text:")) {
      results[results.length - 1] += token.slice("text:".length);
    }
  }
  return results.filter((text) => text !== "").join(" ");
};

const resultOf = (field: ComplexField): string =>
  field.fieldResult
    .flatMap((run) => run.content.flatMap((piece) => (piece.type === "text" ? [piece.text] : [])))
    .join("");

/**
 * The field results a paragraph shows, as one line: what its marker shows
 * after its own text, then each field that stands on the line. A capture
 * shows nothing, so a field hidden behind a marker that does not show it is
 * missing here and present in the file.
 */
export const fieldResultsShown = (paragraph: Paragraph): string => {
  const marker = paragraph.listRendering?.marker ?? "";
  const tab = marker.indexOf("\t");
  const shown = tab === -1 || paragraph.listRendering?.isBullet ? [] : [marker.slice(tab + 1)];
  for (const item of paragraph.content) {
    if (item.type === "complexField") {
      shown.push(resultOf(item));
    }
  }
  return shown.filter((text) => text !== "").join(" ");
};

/**
 * Every paragraph of a live document is in the form the fold allows: its
 * captures open it, and its marker shows exactly their fields.
 */
export const liveFoldFaults = (doc: PMNode): string[] => {
  const faults: string[] = [];
  doc.descendants((node) => {
    if (node.type.name !== "paragraph") {
      return true;
    }
    const marker: unknown = node.attrs["listMarker"];
    const tab = typeof marker === "string" ? marker.indexOf("\t") : -1;
    const suffix = typeof marker === "string" && tab !== -1 ? marker.slice(tab + 1) : "";
    const hidden: string[] = [];
    let opening = true;
    // oxlint-disable-next-line unicorn/no-array-for-each -- ProseMirror Node.forEach
    node.forEach((child) => {
      const folded: unknown =
        child.type.name === "preservedXml" ? child.attrs["foldedListNumber"] : null;
      if (isFoldedListNumber(folded)) {
        if (!opening) {
          faults.push(`a capture stands behind shown content in ${String(node.attrs["paraId"])}`);
        }
        if (folded.kind === "field") {
          hidden.push(resultOf(folded.field));
        }
        return;
      }
      if (child.isText || child.type.name === "field" || child.type.name === "tab") {
        opening = false;
      }
    });
    const cached = hidden.filter((text) => text !== "").join(" ");
    if (cached !== suffix) {
      faults.push(
        `${String(node.attrs["paraId"])} hides "${cached}" and its marker shows "${suffix}"`,
      );
    }
    return true;
  });
  return faults;
};

export const openDocx = (buffer: ArrayBuffer): Promise<Document> =>
  parseDocx(buffer, { preloadFonts: false, detectVariables: false });

export const bodyParagraphs = (model: Document): Paragraph[] =>
  model.package.document.content.filter((block): block is Paragraph => block.type === "paragraph");

/** The marker each paragraph is laid out with. */
export const layoutMarkers = (model: Document): (string | null)[] =>
  toFlowBlocks(toProseDoc(model))
    .filter((block) => block.kind === "paragraph")
    .map((block) => block.attrs?.listMarker ?? null);

/** The marker each paragraph's model carries. */
export const modelMarkers = (model: Document): (string | null)[] =>
  bodyParagraphs(model).map((paragraph) => paragraph.listRendering?.marker ?? null);

/**
 * What each paragraph's content holds, by kind and in order, with the fold's
 * captures named. Neighbouring runs count once: where text is cut into runs
 * is not what the fold decides.
 */
export const contentShapes = (model: Document): string[][] =>
  bodyParagraphs(model).map((paragraph) => {
    const shape: string[] = [];
    for (const item of paragraph.content) {
      const folded = foldedListNumberOf(item);
      const kind = folded ? `folded:${folded.kind}` : item.type;
      if (kind !== "run" || shape.at(-1) !== "run") {
        shape.push(kind);
      }
    }
    return shape;
  });

/** The fold's captures in a ProseMirror document, in document order. */
export const foldedCaptureNodes = (doc: PMNode): { node: PMNode; position: number }[] => {
  const found: { node: PMNode; position: number }[] = [];
  doc.descendants((node, position) => {
    if (node.type.name === "preservedXml" && node.attrs["foldedListNumber"] != null) {
      found.push({ node, position });
    }
  });
  return found;
};

export const paragraphNode = (doc: PMNode, paraId: string): { node: PMNode; position: number } => {
  const found: { node: PMNode; position: number }[] = [];
  doc.descendants((node, position) => {
    if (node.type.name === "paragraph" && node.attrs["paraId"] === paraId) {
      found.push({ node, position });
    }
    return found.length === 0;
  });
  const first = found.at(0);
  if (first === undefined) {
    throw new Error(`The editor holds no paragraph ${paraId}`);
  }
  return first;
};

/** Where `text` starts inside the paragraph `paraId`. */
export const positionOfText = (doc: PMNode, paraId: string, text: string): number => {
  const { node, position } = paragraphNode(doc, paraId);
  const found: number[] = [];
  node.descendants((child, offset) => {
    const index = child.isText ? (child.text?.indexOf(text) ?? -1) : -1;
    if (index >= 0) {
      found.push(position + 1 + offset + index);
    }
  });
  const first = found.at(0);
  if (first === undefined) {
    throw new Error(`Paragraph ${paraId} holds no text ${text}`);
  }
  return first;
};

/** Type `inserted` after the first character of `text` in the paragraph `paraId`. */
export const typeInto = (
  state: EditorState,
  paraId: string,
  text: string,
  inserted: string,
): EditorState =>
  state.apply(state.tr.insertText(inserted, positionOfText(state.doc, paraId, text) + 1));

/** What {@link typeInto} leaves of `text`. */
export const typedInto = (text: string, inserted: string): string =>
  `${text.slice(0, 1)}${inserted}${text.slice(1)}`;

/** Save as the editor does: patch the changed paragraphs, or rewrite the part. */
export const saveDocx = async (
  document: Document,
  original: ArrayBuffer,
  changedParaIds: readonly string[],
): Promise<ArrayBuffer> =>
  (await attemptSelectiveSave(document, original, {
    changedParaIds: new Set(changedParaIds),
    structuralChange: false,
    hasUntrackedChanges: false,
  })) ?? repackDocx(document, { updateModifiedDate: false });
