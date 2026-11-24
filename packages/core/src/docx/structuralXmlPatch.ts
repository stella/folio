/** Conservative body-paragraph splices; source offsets never come from reserialization. */
import { resolveNamespaceUri } from "@stll/docx-core/xml/namespaces";

import type { StructuralSpliceDeclineReason } from "../types/selectiveSaveDecline";
import { canonicalJson } from "../utils/canonicalJson";
import {
  isPlainDocumentMetadataParagraph,
  parseDocumentBackground,
  parseDocumentBodyTree,
  parseFinalBodySectionProperties,
} from "./documentParser";
import { paraIdAttribute } from "./paraIdAttribute";
import { standalonePreviewLedger } from "./previewBudget";
import {
  spliceSupportsRootNamespaces,
  spliceXml,
  type XmlSplice,
} from "./selectiveXmlPatch";
import {
  materializeReplayFragment,
  replayDestinationOf,
  type ReplayFragmentValidation,
  replayIgnorablePrefixes,
} from "./storyBlockReplay";
import {
  findAttributeByNamespaceUri,
  findChild,
  getAttribute,
  getAttributeByNamespaceUri,
  getChildElements,
  getLocalName,
  getNamespaceUri,
  NAMESPACES,
  parseXmlDocument,
  WORDPROCESSINGML_NAMESPACE_URIS,
  type XmlElement,
} from "./xmlParser";

const isWordElement = (element: XmlElement, name: string): boolean =>
  getLocalName(element.name) === name &&
  WORDPROCESSINGML_NAMESPACE_URIS.has(getNamespaceUri(element) ?? "");

type BodyBlock = { element: XmlElement; start: number; end: number };

/** Match XML tokens, skipping quoted delimiters, comments, CDATA and processing instructions. */
const XML_TOKEN =
  /<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>|<\?[\s\S]*?\?>|<(?:"[^"]*"|'[^']*'|[^'">])*>/gu;

const readBody = (xml: string) => {
  const root = parseXmlDocument(xml);
  if (!root || !isWordElement(root, "document")) return null;
  // Strict conversion belongs to full repack; a Transitional fragment cannot
  // inherit a Strict namespace, even when the producer spells its prefix `w`.
  if (getNamespaceUri(root) !== NAMESPACES.w) return null;
  const body = getChildElements(root).find((child) =>
    isWordElement(child, "body"),
  );
  if (!body) return null;
  const children = getChildElements(body);
  const blocks: BodyBlock[] = [];
  const stack: string[] = [];
  let bodyDepth = -1;
  let bodyEnd = -1;
  let start = -1;
  for (const token of xml.matchAll(XML_TOKEN)) {
    const tag = token[0];
    if (tag.startsWith("<!") || tag.startsWith("<?")) continue;
    const closing = tag.startsWith("</");
    const name = tag
      .slice(closing ? 2 : 1)
      .split(/[\s/>]/u)
      .at(0);
    if (!name) return null;
    if (closing) {
      if (stack.pop() !== name) return null;
      if (stack.length === bodyDepth) {
        const element = children[blocks.length];
        if (!element || start < 0) return null;
        blocks.push({ element, start, end: token.index + tag.length });
        start = -1;
      }
      if (name === body.name && stack.length === bodyDepth - 1)
        bodyEnd = token.index;
      continue;
    }
    if (name === body.name && stack.length === 1) bodyDepth = 2;
    if (stack.length === bodyDepth) {
      if (children[blocks.length]?.name !== name) return null;
      start = token.index;
      if (tag.endsWith("/>")) {
        const element = children[blocks.length];
        if (!element) return null;
        blocks.push({ element, start, end: token.index + tag.length });
        start = -1;
      }
    }
    if (!tag.endsWith("/>")) stack.push(name);
  }
  if (stack.length > 0 || blocks.length !== children.length || bodyEnd < 0)
    return null;
  return { root, body, blocks, bodyEnd };
};

function* descendants(element: XmlElement): Generator<XmlElement> {
  yield element;
  for (const child of getChildElements(element)) yield* descendants(child);
}

/**
 * Markers of ranges and fields that can open in one paragraph and close in
 * another: bookmarks, comments, permissions, tracked moves and complex field
 * characters. A splice keeps the source bytes of every untouched block and
 * writes the serialized XML of the others, so it holds exactly when the
 * spliced body places these markers where the serialized body does.
 */
const RANGE_MARKERS = new Set([
  "commentRangeStart",
  "commentRangeEnd",
  "commentReference",
  "bookmarkStart",
  "bookmarkEnd",
  "permStart",
  "permEnd",
  "fldChar",
  "moveFromRangeStart",
  "moveFromRangeEnd",
  "moveToRangeStart",
  "moveToRangeEnd",
]);
/**
 * Range markers the model keeps as body blocks between paragraphs. Each is a
 * barrier like a table: kept as source bytes while the model's marker is
 * unchanged.
 */
const BODY_MARKERS = new Set([
  "bookmarkStart",
  "bookmarkEnd",
  "commentRangeStart",
  "commentRangeEnd",
  "moveFromRangeStart",
  "moveFromRangeEnd",
  "moveToRangeStart",
  "moveToRangeEnd",
]);
const DEPENDENT_ELEMENTS = new Set([
  "sectPr",
  "footnoteReference",
  "endnoteReference",
  "drawing",
  "pict",
  "object",
]);

const paragraphIsSafe = (element: XmlElement, touched: boolean): boolean => {
  for (const child of descendants(element)) {
    if (!WORDPROCESSINGML_NAMESPACE_URIS.has(getNamespaceUri(child) ?? ""))
      continue;
    const name = getLocalName(child.name);
    if (touched && DEPENDENT_ELEMENTS.has(name)) return false;
  }
  return true;
};

const paragraphIds = (root: XmlElement): Set<string> | null => {
  const ids = new Set<string>();
  for (const element of descendants(root)) {
    if (!isWordElement(element, "p")) continue;
    const id = paraIdAttribute(element);
    if (
      !id ||
      !/^[0-9a-f]{8}$/iu.test(id) ||
      /^0+$/u.test(id) ||
      ids.has(id.toUpperCase())
    ) {
      return null;
    }
    ids.add(id.toUpperCase());
  }
  return ids;
};

/** Each body block by paragraph id, or by its position among the barriers. */
const indexBlocks = (blocks: readonly BodyBlock[]) => {
  let barrier = 0;
  const indexed = new Map<string, BodyBlock>();
  for (const block of blocks) {
    const key = isWordElement(block.element, "p")
      ? `p:${paraIdAttribute(block.element)?.toUpperCase()}`
      : `barrier:${barrier++}`;
    indexed.set(key, block);
  }
  return indexed;
};
const barrierOf = (key: string) =>
  key.startsWith("barrier:") ? Number(key.slice("barrier:".length)) : -1;

/**
 * An attribute as the reader takes it: in a WordprocessingML namespace, else
 * unprefixed, with a decimal number read as its value.
 */
const markerAttribute = (element: XmlElement, name: string): string => {
  const value =
    findAttributeByNamespaceUri(element, WORDPROCESSINGML_NAMESPACE_URIS, name)
      ?.value ?? getAttribute(element, "w", name);
  if (value === null) return "";
  const trimmed = value.trim();
  return /^-?\d+$/u.test(trimmed) ? String(Number(trimmed)) : value;
};

/** One range or field marker by kind and identity. */
const rangeMarker = (element: XmlElement): string => {
  const kind = getLocalName(element.name);
  if (kind === "fldChar")
    return `fldChar:${markerAttribute(element, "fldCharType")}`;
  const name = markerAttribute(element, "name");
  const id = markerAttribute(element, "id");
  return name === "" ? `${kind}:${id}` : `${kind}:${id}:${name}`;
};

/** Each body block with the range and field markers it holds, in order. */
const rangeMarkers = (blocks: readonly BodyBlock[]): string[][] =>
  [...indexBlocks(blocks)].map(([key, { element }]) => {
    const markers = [key];
    for (const child of descendants(element))
      if (
        WORDPROCESSINGML_NAMESPACE_URIS.has(getNamespaceUri(child) ?? "") &&
        RANGE_MARKERS.has(getLocalName(child.name))
      )
        markers.push(rangeMarker(child));
    return markers;
  });

type StructuralPatchOptions = {
  originalXml: string;
  serializedXml: string;
  changedIds: ReadonlySet<string>;
  fragmentValidation?: ReplayFragmentValidation | undefined;
  /** Demand proof that the authored body can receive raw generated paragraphs. */
  emptySourceParagraphCount?: number;
};

/** The spliced part, or why the splice declined and the save repacks. */
export type StructuralPatchResult =
  | { type: "patched"; xml: string }
  | { type: "declined"; reason: StructuralSpliceDeclineReason };

/** Verbatim captures: they keep the source's own spelling, not its meaning. */
const CAPTURE_KEYS: ReadonlySet<string> = new Set([
  "sourceXml",
  "gridSourceXml",
]);

/** `value` without its verbatim captures, so a respelled source compares equal. */
const withoutCaptures = (value: unknown): unknown => {
  if (Array.isArray(value)) return value.map(withoutCaptures);
  if (typeof value !== "object" || value === null) return value;
  return Object.fromEntries(
    Object.entries(value)
      .filter(([key]) => !CAPTURE_KEYS.has(key))
      .map(([key, field]) => [key, withoutCaptures(field)]),
  );
};

const decline = (
  reason: StructuralSpliceDeclineReason,
): StructuralPatchResult => ({ type: "declined", reason });

const isBodyMarker = (element: XmlElement): boolean =>
  WORDPROCESSINGML_NAMESPACE_URIS.has(getNamespaceUri(element) ?? "") &&
  BODY_MARKERS.has(getLocalName(element.name));

type ReadBody = NonNullable<ReturnType<typeof readBody>>;
type EmptySourceShapeOptions = {
  source: ReadBody;
  current: ReadBody;
  paragraphCount: number;
};
const MC_NAMESPACES = new Set([NAMESPACES.mc]);

const coversEmptySource = ({
  source,
  current,
  paragraphCount,
}: EmptySourceShapeOptions): boolean => {
  const beforeSection = source.blocks.at(0)?.element;
  const afterSection = current.blocks.at(-1)?.element;
  if (
    paragraphCount < 1 ||
    source.blocks.length !== 1 ||
    current.blocks.length !== paragraphCount + 1 ||
    beforeSection === undefined ||
    afterSection === undefined ||
    !isWordElement(beforeSection, "sectPr") ||
    !isWordElement(afterSection, "sectPr") ||
    getNamespaceUri(source.body) !== NAMESPACES.w ||
    getNamespaceUri(current.body) !== NAMESPACES.w ||
    getChildElements(source.root).length !== 1 ||
    getChildElements(current.root).length !== 1 ||
    current.blocks
      .slice(0, -1)
      .some(({ element }) => !isWordElement(element, "p")) ||
    current.body.elements?.some(
      (node) => node.type === "text" && String(node.text ?? "").trim() !== "",
    )
  )
    return false;

  // Replay binds the generated body scope on every fragment. When that
  // fragment is read back inside the source, it must inherit no extra
  // declarations. A second materialization then adds none.
  for (let scope = source.body.namespaceScope; scope; scope = scope.parent)
    for (const [prefix, uri] of scope.bindings)
      if (resolveNamespaceUri(current.body.namespaceScope, prefix) !== uri)
        return false;

  // The second materialization prepends the source tokens to the already
  // materialized generated tokens. Their authored order cannot change.
  const ignorable = (root: XmlElement) =>
    replayIgnorablePrefixes(
      getAttributeByNamespaceUri(root, MC_NAMESPACES, "Ignorable") ?? "",
    );
  const generated = ignorable(current.root);
  return ignorable(source.root).every(
    (prefix, index) => prefix === generated[index],
  );
};

const parseBodyModel = (root: XmlElement) =>
  parseDocumentBodyTree({
    doc: { elements: [root] },
    styles: null,
    theme: null,
    numbering: null,
    rels: null,
    media: null,
    previews: standalonePreviewLedger(),
  });

/**
 * Insert/delete direct body paragraphs between surviving paragraph/table anchors.
 * Tables and body-level range markers are opaque barriers: their modeled
 * content and order must stay identical. Ranges and fields may cross the
 * spliced paragraphs: the result must place their markers exactly where the
 * serialized body does. Id-less sources need an explicit ensureParaIds ingest
 * before structural editing.
 */
export const buildStructuralDocumentPatch = ({
  originalXml,
  serializedXml,
  changedIds,
  fragmentValidation,
  emptySourceParagraphCount,
}: StructuralPatchOptions): StructuralPatchResult => {
  const source = readBody(originalXml);
  const current = readBody(serializedXml);
  if (!source || !current) return decline("structuralUnreadableBody");
  const allSourceIds = paragraphIds(source.root);
  const allCurrentIds = paragraphIds(current.root);
  if (!allSourceIds || !allCurrentIds) return decline("structuralParagraphIds");
  const changed = new Set([...changedIds].map((id) => id.toUpperCase()));
  let plainBody =
    findChild(source.root, "w", "body") === source.body &&
    findChild(current.root, "w", "body") === current.body;
  for (const { element } of [...source.blocks, ...current.blocks]) {
    if (isWordElement(element, "p")) {
      const id = paraIdAttribute(element)?.toUpperCase();
      if (!id) return decline("structuralParagraphIds");
      if (plainBody && isPlainDocumentMetadataParagraph(element)) continue;
      plainBody = false;
      const touched =
        changed.has(id) || !allSourceIds.has(id) || !allCurrentIds.has(id);
      if (!paragraphIsSafe(element, touched))
        return decline("structuralDependentContent");
    } else if (!isWordElement(element, "sectPr")) {
      plainBody = false;
      if (!isWordElement(element, "tbl") && !isBodyMarker(element))
        return decline("structuralBodyBlock");
    }
  }
  if (
    emptySourceParagraphCount !== undefined &&
    (!plainBody ||
      allSourceIds.size !== 0 ||
      allCurrentIds.size !== emptySourceParagraphCount ||
      !coversEmptySource({
        source,
        current,
        paragraphCount: emptySourceParagraphCount,
      }))
  )
    return decline("structuralBodyBlock");
  // Barriers are matched by position, so the source's must be the same
  // elements as the serialized ones: a marker the reader dropped would shift
  // every barrier after it.
  const barrierNames = (blocks: readonly BodyBlock[]) =>
    blocks
      .filter(({ element }) => !isWordElement(element, "p"))
      .map(({ element }) => getLocalName(element.name));
  if (
    canonicalJson(barrierNames(source.blocks)) !==
    canonicalJson(barrierNames(current.blocks))
  )
    return decline("structuralBarrierChanged");

  // Compare both sides through the same parser, avoiding lexical differences
  // such as empty tags and attribute order. Unmodeled table bytes stay in source.
  const barrierBody = ({
    root,
    body,
  }: NonNullable<ReturnType<typeof readBody>>) =>
    plainBody
      ? {
          content: [],
          finalSectionProperties: parseFinalBodySectionProperties(body),
          background: parseDocumentBackground(root, undefined),
        }
      : parseBodyModel(root);
  const sourceBody = barrierBody(source);
  const currentBody = barrierBody(current);
  const barriers = (body: typeof sourceBody) =>
    body.content.filter((block) => block.type !== "paragraph");
  const sourceBarriers = barriers(sourceBody);
  const currentBarriers = barriers(currentBody);
  if (
    sourceBarriers.length !== currentBarriers.length ||
    canonicalJson({
      finalSectionProperties: sourceBody.finalSectionProperties,
      background: sourceBody.background,
    }) !==
      canonicalJson({
        finalSectionProperties: currentBody.finalSectionProperties,
        background: currentBody.background,
      })
  )
    return decline("structuralBarrierChanged");
  // A table whose content changed is written whole from the serialized part;
  // every other table keeps its source bytes, including what the model does
  // not represent. Captures are left out of the comparison: they record how
  // the source spelled its markup, which a serialized part need not repeat.
  const changedTables = new Set<number>();
  for (const [index, block] of sourceBarriers.entries()) {
    const next = currentBarriers[index];
    if (
      canonicalJson(withoutCaptures(block)) ===
      canonicalJson(withoutCaptures(next))
    )
      continue;
    if (block.type !== "table" || next?.type !== "table")
      return decline("structuralBarrierChanged");
    changedTables.add(index);
  }

  const before = indexBlocks(source.blocks);
  const after = indexBlocks(current.blocks);
  const survivingBefore = [...before.keys()].filter((key) => after.has(key));
  const survivingAfter = [...after.keys()].filter((key) => before.has(key));
  if (canonicalJson(survivingBefore) !== canonicalJson(survivingAfter))
    return decline("structuralReorder");

  const destination = replayDestinationOf(source.root, source.body);
  const fragmentFor = (block: BodyBlock) =>
    materializeReplayFragment({
      xml: serializedXml.slice(block.start, block.end),
      element: block.element,
      sourceNamespace: getNamespaceUri(source.root) ?? "",
      generatedRoot: current.root,
      validation: fragmentValidation,
      destination,
    });
  const splices: XmlSplice[] = [];
  for (const [key, block] of before) {
    if (!after.has(key))
      splices.push({ start: block.start, end: block.end, newXml: "" });
  }
  let pending: string[] = [];
  for (const [key, block] of after) {
    const original = before.get(key);
    if (!original) {
      const id = paraIdAttribute(block.element)?.toUpperCase();
      if (!id || allSourceIds.has(id)) return decline("structuralReorder");
      const fragment = fragmentFor(block);
      if (fragment === null) return decline("structuralFragment");
      pending.push(fragment);
      continue;
    }
    const id = paraIdAttribute(block.element)?.toUpperCase();
    const replacement =
      (id !== undefined && changed.has(id)) || changedTables.has(barrierOf(key))
        ? fragmentFor(block)
        : originalXml.slice(original.start, original.end);
    if (replacement === null) return decline("structuralFragment");
    if (
      pending.length > 0 ||
      replacement !== originalXml.slice(original.start, original.end)
    ) {
      splices.push({
        start: original.start,
        end: original.end,
        newXml: pending.join("") + replacement,
      });
    }
    pending = [];
  }
  if (pending.length > 0)
    splices.push({
      start: source.bodyEnd,
      end: source.bodyEnd,
      newXml: pending.join(""),
    });
  if (
    !spliceSupportsRootNamespaces({
      originalXml,
      serializedXml,
      fragments: splices.map(({ newXml }) => newXml).filter(Boolean),
      originalRoot: source.root,
      serializedRoot: current.root,
    })
  )
    return decline("structuralFragment");
  const patched = spliceXml(originalXml, splices);
  if (patched === null) return decline("structuralSpliceRefused");
  // Untouched blocks keep their source bytes, so the range and field markers
  // in them are the source's. The splice holds only where those, with the
  // ones the serialized blocks carry, are exactly the serialized body's, block
  // by block and in order: no marker dropped, duplicated, left behind or
  // moved to another block, so no range or field is split by the splice.
  const spliced = readBody(patched);
  if (!spliced) return decline("structuralSpliceRefused");
  if (
    canonicalJson(rangeMarkers(spliced.blocks)) !==
    canonicalJson(rangeMarkers(current.blocks))
  )
    return decline("structuralRangeSplit");
  return { type: "patched", xml: patched };
};
