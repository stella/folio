import { escapeXmlAttribute } from "@stll/docx-core";
import { DOCX_CONFORMANCE_CLASSES } from "@stll/docx-core/model";
import {
  getXmlSourceRange,
  parseStreamingXmlWithSourceRanges,
} from "@stll/docx-core/xml";

import type { BlockContent, DocumentBody } from "../types/document";
import { canonicalJson } from "../utils/canonicalJson";
import {
  type SerializedPart,
  spliceXml,
  type XmlSplice,
} from "./selectiveXmlPatch";
import { TRANSITIONAL_NAMESPACE_BY_STRICT_URI } from "./strictValueEncodings.gen";
import {
  captureSourceProfileXml,
  isSafeCapturedXmlDocument,
} from "./verbatimCapture";
import {
  getAttributeByNamespaceUri,
  getChildElements,
  getLocalName,
  getNamespaceUri,
  WORDPROCESSINGML_NAMESPACE_URIS,
  type XmlElement,
  type XmlNamespaceScope,
} from "./xmlParser";

/**
 * What a block saves as, for telling whether it changed: the block without
 * the list marker a paragraph's numbering computes, which no save writes. An
 * edit elsewhere in the list renumbers it; the paragraph itself is unchanged
 * and keeps its source.
 */
const savedFingerprint = (value: unknown): string => {
  if (value === null || value === undefined) return "null";
  if (Array.isArray(value)) return `[${value.map(savedFingerprint).join(",")}]`;
  if (typeof value !== "object") return JSON.stringify(value) ?? "null";
  const paragraph = "type" in value && value.type === "paragraph";
  const entries = Object.entries(value)
    .filter(
      ([key, entry]) =>
        entry !== undefined && !(paragraph && key === "listRendering"),
    )
    .toSorted(([left], [right]) => left.localeCompare(right));
  return `{${entries
    .map(([key, entry]) => `${JSON.stringify(key)}:${savedFingerprint(entry)}`)
    .join(",")}}`;
};

const BLOCK_ELEMENT_NAMES = {
  paragraph: "p",
  table: "tbl",
  blockSdt: "sdt",
  blockCustomXml: "customXml",
  preservedBlock: null,
  bookmarkStart: "bookmarkStart",
  bookmarkEnd: "bookmarkEnd",
  commentRangeStart: "commentRangeStart",
  commentRangeEnd: "commentRangeEnd",
  moveFromRangeStart: "moveFromRangeStart",
  moveFromRangeEnd: "moveFromRangeEnd",
  moveToRangeStart: "moveToRangeStart",
  moveToRangeEnd: "moveToRangeEnd",
} as const satisfies Record<BlockContent["type"], string | null>;
type ModelMatchOptions = {
  value: BlockContent | undefined;
  element: XmlElement;
  mode: "source" | "generated";
};
const modelMatchesElement = ({
  value,
  element,
  mode,
}: ModelMatchOptions): boolean => {
  if (value === undefined) return false;
  // The parser projects a supported AlternateContent branch as one block;
  // source cardinality and the trusted baseline preserve its outer carrier.
  if (
    mode === "source" &&
    getLocalName(element.name) === "AlternateContent" &&
    getNamespaceUri(element) ===
      "http://schemas.openxmlformats.org/markup-compatibility/2006" &&
    (value.type === "paragraph" || value.type === "table")
  )
    return true;
  const name = BLOCK_ELEMENT_NAMES[value.type];
  if (name !== null)
    return (
      getLocalName(element.name) === name &&
      WORDPROCESSINGML_NAMESPACE_URIS.has(getNamespaceUri(element) ?? "")
    );
  if (value.type !== "preservedBlock") return false;
  const parsed = parseStreamingXmlWithSourceRanges(
    value.xml,
    element.namespaceScope,
  );
  if (parsed.status !== "parsed") return false;
  const roots = getChildElements(parsed.value);
  const root = roots.at(0);
  const normalize = (uri: string | undefined) =>
    uri === undefined
      ? undefined
      : (TRANSITIONAL_NAMESPACE_BY_STRICT_URI.get(uri) ?? uri);
  return (
    roots.length === 1 &&
    root !== undefined &&
    getLocalName(root.name) === getLocalName(element.name) &&
    normalize(getNamespaceUri(root)) === normalize(getNamespaceUri(element))
  );
};
// Each complete XML token includes quoted delimiters; text, comments and CDATA
// never participate in namespace rewriting.
const XML_TOKEN =
  /<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>|<\?[\s\S]*?\?>|<(?:"[^"]*"|'[^']*'|[^'">])*>/gu;
const XML_ATTRIBUTE = /([^\s=<>/]+)(\s*=\s*)(["'])(.*?)\3/gsu;

type XmlRange = { readonly start: number; readonly end: number };
type StoryBlock = {
  readonly element: XmlElement;
  readonly start: number;
  readonly end: number;
};
/** A story element and its block children; offsets index the XML it was read from. */
type ReadStory = {
  readonly root: XmlElement;
  readonly blocks: readonly StoryBlock[];
  readonly range: XmlRange;
};

const storyOfElement = (
  root: XmlElement,
  ignoredChild?: string,
): ReadStory | null => {
  if (!WORDPROCESSINGML_NAMESPACE_URIS.has(getNamespaceUri(root) ?? ""))
    return null;
  if (
    root.elements?.some(
      (child) =>
        child.type === "text" && String(child.text ?? "").trim() !== "",
    )
  )
    return null;
  const children = getChildElements(root);
  const blocks = [];
  for (const element of children) {
    if (
      getLocalName(element.name) === ignoredChild &&
      WORDPROCESSINGML_NAMESPACE_URIS.has(getNamespaceUri(element) ?? "")
    )
      continue;
    const range = getXmlSourceRange(element);
    if (!range) return null;
    blocks.push({ element, ...range });
  }
  const range = getXmlSourceRange(root);
  return range ? { root, blocks, range } : null;
};

const readStory = (
  xml: string,
  scope?: XmlNamespaceScope,
  ignoredChild?: string,
): ReadStory | null => {
  const parsed = parseStreamingXmlWithSourceRanges(xml, scope);
  if (parsed.status !== "parsed") return null;
  const roots = getChildElements(parsed.value);
  const root = roots.at(0);
  if (roots.length !== 1 || !root) return null;
  return storyOfElement(root, ignoredChild);
};

const effectiveBindings = (element: XmlElement): Map<string, string> => {
  const bindings = new Map<string, string>();
  let scope = element.namespaceScope;
  while (scope) {
    for (const [prefix, uri] of scope.bindings)
      if (!bindings.has(prefix)) bindings.set(prefix, uri);
    scope = scope.parent;
  }
  return bindings;
};

type MaterializeReplayFragmentOptions = {
  xml: string;
  element: XmlElement;
  sourceNamespace: string;
  generatedRoot?: XmlElement;
};

/** Bind a replay fragment once, retaining existing local declarations and compatibility tokens. */
export const materializeReplayFragment = ({
  xml,
  element,
  sourceNamespace,
  generatedRoot,
}: MaterializeReplayFragmentOptions): string | null => {
  const strict = TRANSITIONAL_NAMESPACE_BY_STRICT_URI.has(sourceNamespace);
  const tokens = [...xml.matchAll(XML_TOKEN)];
  const first = tokens.at(0);
  if (
    !first ||
    first.index !== 0 ||
    first[0].startsWith("<!") ||
    first[0].startsWith("<?")
  )
    return null;
  const opening = first[0];
  const additions = [];
  const bindings = effectiveBindings(element);
  const markupCompatibilityUri =
    "http://schemas.openxmlformats.org/markup-compatibility/2006";
  const ignorableUris = new Set([markupCompatibilityUri]);
  const inheritedIgnorable =
    generatedRoot === undefined
      ? null
      : getAttributeByNamespaceUri(generatedRoot, ignorableUris, "Ignorable");
  const localIgnorable = getAttributeByNamespaceUri(
    element,
    ignorableUris,
    "Ignorable",
  );
  const ignorable = [
    ...new Set(
      `${inheritedIgnorable ?? ""} ${localIgnorable ?? ""}`
        .trim()
        .split(/\s+/u)
        .filter(Boolean),
    ),
  ];
  let ignorableAttribute: string | undefined;
  if (ignorable.length > 0) {
    const localAttribute = Object.keys(element.attributes ?? {}).find(
      (name) => {
        const colon = name.indexOf(":");
        return (
          colon > 0 &&
          getLocalName(name) === "Ignorable" &&
          bindings.get(name.slice(0, colon)) === markupCompatibilityUri
        );
      },
    );
    const compatibilityPrefix =
      localAttribute?.slice(0, localAttribute.indexOf(":")) ??
      [...bindings].find(
        ([prefix, uri]) => prefix !== "" && uri === markupCompatibilityUri,
      )?.[0];
    if (
      compatibilityPrefix === undefined ||
      ignorable.some((prefix) => !bindings.has(prefix))
    )
      return null;
    ignorableAttribute = `${compatibilityPrefix}:Ignorable`;
    if (localIgnorable === null)
      additions.push(
        ` ${ignorableAttribute}="${escapeXmlAttribute(ignorable.join(" "))}"`,
      );
  }
  for (const [prefix, uri] of bindings) {
    const name = prefix === "" ? "xmlns" : `xmlns:${prefix}`;
    if (element.attributes?.[name] !== undefined) continue;
    additions.push(` ${name}="${escapeXmlAttribute(uri)}"`);
  }
  let replacedOpening = opening.replace(
    XML_ATTRIBUTE,
    (match, name: string, assignment: string, quote: string) =>
      name === ignorableAttribute
        ? `${name}${assignment}${quote}${escapeXmlAttribute(ignorable.join(" "))}${quote}`
        : match,
  );
  const close = opening.endsWith("/>")
    ? replacedOpening.length - 2
    : replacedOpening.length - 1;
  replacedOpening =
    replacedOpening.slice(0, close) +
    additions.join("") +
    replacedOpening.slice(close);
  const materialized = replacedOpening + xml.slice(opening.length);
  if (!isSafeCapturedXmlDocument(materialized)) return null;
  if (!strict) return materialized;
  const parsed = parseStreamingXmlWithSourceRanges(materialized);
  if (parsed.status !== "parsed") return null;
  const root = getChildElements(parsed.value).at(0);
  return root === undefined
    ? null
    : captureSourceProfileXml(root, DOCX_CONFORMANCE_CLASSES.STRICT);
};

const paragraphIdentity = (value: BlockContent): string | undefined =>
  value.type === "paragraph" ? value.paraId?.toUpperCase() : undefined;

type CandidateQueue = { indices: number[]; cursor: number };
type AddCandidateOptions = {
  map: Map<string, CandidateQueue>;
  key: string;
  index: number;
};
const addCandidate = ({ map, key, index }: AddCandidateOptions): void => {
  const queue = map.get(key);
  if (queue) queue.indices.push(index);
  else map.set(key, { indices: [index], cursor: 0 });
};
const takeCandidate = (
  queue: CandidateQueue | undefined,
  used: ReadonlySet<number>,
): number | undefined => {
  if (!queue) return undefined;
  while (queue.cursor < queue.indices.length) {
    const candidate = queue.indices[queue.cursor];
    queue.cursor += 1;
    if (candidate !== undefined && !used.has(candidate)) return candidate;
  }
  return undefined;
};

type ReplayScopeOptions = {
  source: XmlNamespaceScope | undefined;
  generated: XmlNamespaceScope | undefined;
  generatedRoot: XmlElement | undefined;
};

type StoryBlockReplayOptions = {
  sourceXml: string;
  baselineContent: readonly BlockContent[];
  currentContent: readonly BlockContent[];
  serializedXml: string;
  ignoredChild?: string;
};

/**
 * A source story whose blocks each match their baseline model block, as a
 * replay reads it: block ranges and the baseline's saved fingerprints. Only
 * container blocks keep their element, for a nested replay; a source read once
 * serves every later save of the same part without holding its tree.
 */
type SourceStory = {
  readonly rootName: string | undefined;
  readonly namespace: string;
  readonly range: XmlRange;
  readonly blocks: readonly {
    readonly start: number;
    readonly end: number;
    readonly element: XmlElement | undefined;
  }[];
  /** Start of the ignored child, where blocks enter a story that has none. */
  readonly ignoredStart: number | undefined;
  readonly fingerprints: readonly string[];
};

const isContainerBlock = (block: BlockContent | undefined): boolean =>
  block?.type === "table" ||
  block?.type === "blockSdt" ||
  block?.type === "blockCustomXml";

const readSourceStory = (
  story: ReadStory,
  baselineContent: readonly BlockContent[],
  ignoredChild: string | undefined,
): SourceStory | null => {
  if (
    story.blocks.length !== baselineContent.length ||
    story.blocks.some(
      (block, index) =>
        !modelMatchesElement({
          value: baselineContent[index],
          element: block.element,
          mode: "source",
        }),
    )
  )
    return null;
  const ignored =
    ignoredChild === undefined
      ? undefined
      : getChildElements(story.root).find(
          (child) =>
            getLocalName(child.name) === ignoredChild &&
            WORDPROCESSINGML_NAMESPACE_URIS.has(getNamespaceUri(child) ?? ""),
        );
  return {
    rootName: story.root.name,
    namespace: getNamespaceUri(story.root) ?? "",
    range: story.range,
    blocks: story.blocks.map(({ element, start, end }, index) => ({
      start,
      end,
      element: isContainerBlock(baselineContent[index]) ? element : undefined,
    })),
    ignoredStart: ignored && getXmlSourceRange(ignored)?.start,
    fingerprints: baselineContent.map(savedFingerprint),
  };
};

type StoryReplayOptions = {
  /** The XML the source story's offsets index. */
  sourceXml: string;
  /** The XML the generated story's offsets index. */
  serializedXml: string;
  source: SourceStory;
  generated: ReadStory;
  baselineContent: readonly BlockContent[];
  currentContent: readonly BlockContent[];
  generatedRoot: XmlElement;
};

/**
 * Replace only mapped block ranges, retaining authored root syntax and every
 * intervening gap. The splices are in the source story's offsets.
 */
const replayStoryBlocks = ({
  sourceXml,
  serializedXml,
  source,
  generated,
  baselineContent,
  currentContent,
  generatedRoot,
}: StoryReplayOptions): XmlSplice[] | null => {
  if (
    getLocalName(source.rootName) !== getLocalName(generated.root.name) ||
    generated.blocks.length !== currentContent.length ||
    generated.blocks.some(
      (block, index) =>
        !modelMatchesElement({
          value: currentContent[index],
          element: block.element,
          mode: "generated",
        }),
    )
  )
    return null;
  const { fingerprints } = source;
  const identities = new Map<string, CandidateQueue>();
  const byFingerprint = new Map<string, CandidateQueue>();
  for (const [index, block] of baselineContent.entries()) {
    const fingerprint = fingerprints[index];
    if (fingerprint === undefined) return null;
    addCandidate({ map: byFingerprint, key: fingerprint, index });
    const identity = paragraphIdentity(block);
    if (identity !== undefined)
      addCandidate({ map: identities, key: identity, index });
  }
  const used = new Set<number>();
  const fragments: string[] = [];
  for (const [index, block] of currentContent.entries()) {
    const fingerprint = savedFingerprint(block);
    const identity = paragraphIdentity(block);
    let originalIndex =
      takeCandidate(
        identity === undefined ? undefined : identities.get(identity),
        used,
      ) ?? takeCandidate(byFingerprint.get(fingerprint), used);
    // Containers have no durable block id. An unchanged wrapper at the same
    // ordinal can still retain its source while its nested blocks are edited.
    if (
      originalIndex === undefined &&
      (block.type === "table" ||
        block.type === "blockSdt" ||
        block.type === "blockCustomXml") &&
      !used.has(index)
    ) {
      const baseline = baselineContent[index];
      if (baseline?.type === block.type) originalIndex = index;
    }
    if (originalIndex !== undefined) used.add(originalIndex);
    const original =
      originalIndex === undefined ? undefined : source.blocks[originalIndex];
    if (
      originalIndex !== undefined &&
      original &&
      fingerprints[originalIndex] === fingerprint
    ) {
      fragments.push(sourceXml.slice(original.start, original.end));
      continue;
    }
    const replacement = generated.blocks[index];
    if (!replacement) return null;
    if (original && originalIndex !== undefined && original.element) {
      const nested = replayNestedBlock({
        sourceXml,
        serializedXml,
        sourceElement: original.element,
        generatedElement: replacement.element,
        baseline: baselineContent[originalIndex],
        current: block,
        generatedRoot,
      });
      if (nested !== null) {
        fragments.push(nested);
        continue;
      }
    }
    const fragment = materializeReplayFragment({
      xml: serializedXml.slice(replacement.start, replacement.end),
      element: replacement.element,
      sourceNamespace: source.namespace,
      generatedRoot,
    });
    if (fragment === null) return null;
    fragments.push(fragment);
  }
  const splices: XmlSplice[] = source.blocks.map((block, index) => ({
    start: block.start,
    end: block.end,
    newXml:
      index === source.blocks.length - 1
        ? fragments.slice(index).join("")
        : (fragments[index] ?? ""),
  }));
  if (source.blocks.length === 0 && fragments.length > 0) {
    const opening = [
      ...sourceXml
        .slice(source.range.start, source.range.end)
        .matchAll(XML_TOKEN),
    ].at(0)?.[0];
    if (!opening || !source.rootName) return null;
    if (opening.endsWith("/>")) {
      splices.push({
        start: source.range.start,
        end: source.range.end,
        newXml: `${opening.slice(0, -2)}>${fragments.join("")}</${source.rootName}>`,
      });
    } else {
      // The story's own closing tag: the last one that starts inside it.
      const close =
        source.ignoredStart ??
        sourceXml.lastIndexOf("</", source.range.end - 1);
      if (close < source.range.start) return null;
      splices.push({ start: close, end: close, newXml: fragments.join("") });
    }
  }
  return splices;
};

/** Read both stories from their own XML and replay them in one pass. */
const replayBlocks = (
  {
    sourceXml,
    baselineContent,
    currentContent,
    serializedXml,
    ignoredChild,
  }: StoryBlockReplayOptions,
  scope?: ReplayScopeOptions,
): string | null => {
  const sourceStory = readStory(sourceXml, scope?.source, ignoredChild);
  const generated = readStory(serializedXml, scope?.generated, ignoredChild);
  const source =
    sourceStory && readSourceStory(sourceStory, baselineContent, ignoredChild);
  if (!source || !generated) return null;
  const splices = replayStoryBlocks({
    sourceXml,
    serializedXml,
    source,
    generated,
    baselineContent,
    currentContent,
    generatedRoot: scope?.generatedRoot ?? generated.root,
  });
  return splices === null ? null : spliceXml(sourceXml, splices);
};

export const buildStoryBlockReplay = (
  options: StoryBlockReplayOptions,
): string | null => replayBlocks(options);

const withoutChildren = (value: object, key: string) =>
  Object.fromEntries(Object.entries(value).filter(([name]) => name !== key));
const wordChildren = (element: XmlElement, name: string) =>
  getChildElements(element).filter(
    (child) =>
      getLocalName(child.name) === name &&
      WORDPROCESSINGML_NAMESPACE_URIS.has(getNamespaceUri(child) ?? ""),
  );

type NestedBlockOptions = {
  sourceXml: string;
  serializedXml: string;
  sourceElement: XmlElement;
  generatedElement: XmlElement;
  baseline: BlockContent | undefined;
  current: BlockContent;
  generatedRoot: XmlElement;
};
const replayNestedBlock = ({
  sourceXml,
  serializedXml,
  sourceElement,
  generatedElement,
  baseline,
  current,
  generatedRoot,
}: NestedBlockOptions): string | null => {
  if (baseline === undefined || baseline.type !== current.type) return null;
  const sourceRange = getXmlSourceRange(sourceElement);
  if (!sourceRange) return null;
  const splices: XmlSplice[] = [];
  const nestedContent = (
    sourceContainer: XmlElement,
    generatedContainer: XmlElement,
    before: readonly BlockContent[],
    after: readonly BlockContent[],
  ) => {
    const sourceContentRange = getXmlSourceRange(sourceContainer);
    const generatedContentRange = getXmlSourceRange(generatedContainer);
    if (!sourceContentRange || !generatedContentRange) return false;
    const patched = replayBlocks(
      {
        sourceXml: sourceXml.slice(
          sourceContentRange.start,
          sourceContentRange.end,
        ),
        serializedXml: serializedXml.slice(
          generatedContentRange.start,
          generatedContentRange.end,
        ),
        baselineContent: before,
        currentContent: after,
      },
      {
        source: sourceContainer.namespaceScope,
        generated: generatedContainer.namespaceScope,
        generatedRoot,
      },
    );
    if (patched === null) return false;
    splices.push({
      start: sourceContentRange.start - sourceRange.start,
      end: sourceContentRange.end - sourceRange.start,
      newXml: patched,
    });
    return true;
  };
  switch (current.type) {
    case "blockSdt":
    case "blockCustomXml": {
      if (
        (baseline.type !== "blockSdt" && baseline.type !== "blockCustomXml") ||
        canonicalJson(withoutChildren(baseline, "content")) !==
          canonicalJson(withoutChildren(current, "content"))
      )
        return null;
      if (current.type === "blockCustomXml") {
        if (
          !nestedContent(
            sourceElement,
            generatedElement,
            baseline.content,
            current.content,
          )
        )
          return null;
        break;
      }
      const sourceChildren = wordChildren(sourceElement, "sdtContent");
      const generatedChildren = wordChildren(generatedElement, "sdtContent");
      const sourceContainer = sourceChildren.at(0);
      const generatedContainer = generatedChildren.at(0);
      if (
        sourceChildren.length !== 1 ||
        generatedChildren.length !== 1 ||
        !sourceContainer ||
        !generatedContainer ||
        !nestedContent(
          sourceContainer,
          generatedContainer,
          baseline.content,
          current.content,
        )
      )
        return null;
      break;
    }
    case "table": {
      if (
        baseline.type !== "table" ||
        baseline.rows.length !== current.rows.length ||
        canonicalJson(withoutChildren(baseline, "rows")) !==
          canonicalJson(withoutChildren(current, "rows"))
      )
        return null;
      const sourceRows = wordChildren(sourceElement, "tr");
      const generatedRows = wordChildren(generatedElement, "tr");
      if (
        sourceRows.length !== current.rows.length ||
        generatedRows.length !== current.rows.length
      )
        return null;
      for (const [rowIndex, row] of current.rows.entries()) {
        const beforeRow = baseline.rows[rowIndex];
        const sourceRow = sourceRows[rowIndex];
        const generatedRow = generatedRows[rowIndex];
        if (
          beforeRow === undefined ||
          beforeRow.cells.length !== row.cells.length ||
          canonicalJson(withoutChildren(beforeRow, "cells")) !==
            canonicalJson(withoutChildren(row, "cells")) ||
          !sourceRow ||
          !generatedRow
        )
          return null;
        const sourceCells = wordChildren(sourceRow, "tc");
        const generatedCells = wordChildren(generatedRow, "tc");
        if (
          sourceCells.length !== row.cells.length ||
          generatedCells.length !== row.cells.length
        )
          return null;
        for (const [cellIndex, cell] of row.cells.entries()) {
          const beforeCell = beforeRow.cells[cellIndex];
          const sourceCell = sourceCells[cellIndex];
          const generatedCell = generatedCells[cellIndex];
          if (
            beforeCell === undefined ||
            canonicalJson(withoutChildren(beforeCell, "content")) !==
              canonicalJson(withoutChildren(cell, "content")) ||
            !sourceCell ||
            !generatedCell
          )
            return null;
          // tcPr is wrapper metadata; retain it outside the content region.
          if (
            !replayCellContent({
              sourceXml,
              serializedXml,
              sourceCell,
              generatedCell,
              before: beforeCell.content,
              after: cell.content,
              sourceRangeStart: sourceRange.start,
              splices,
              generatedRoot,
            })
          )
            return null;
        }
      }
      break;
    }
    case "paragraph":
    case "preservedBlock":
    case "bookmarkStart":
    case "bookmarkEnd":
    case "commentRangeStart":
    case "commentRangeEnd":
    case "moveFromRangeStart":
    case "moveFromRangeEnd":
    case "moveToRangeStart":
    case "moveToRangeEnd":
      return null;
    default: {
      const unreachable: never = current;
      return unreachable;
    }
  }
  return spliceXml(
    sourceXml.slice(sourceRange.start, sourceRange.end),
    splices,
  );
};

type CellContentOptions = {
  sourceXml: string;
  serializedXml: string;
  sourceCell: XmlElement;
  generatedCell: XmlElement;
  before: readonly BlockContent[];
  after: readonly BlockContent[];
  sourceRangeStart: number;
  splices: XmlSplice[];
  generatedRoot: XmlElement;
};
const replayCellContent = ({
  sourceXml,
  serializedXml,
  sourceCell,
  generatedCell,
  before,
  after,
  sourceRangeStart,
  splices,
  generatedRoot,
}: CellContentOptions): boolean => {
  const sourceRange = getXmlSourceRange(sourceCell);
  const generatedRange = getXmlSourceRange(generatedCell);
  if (!sourceRange || !generatedRange) return false;
  const sourceProperties = wordChildren(sourceCell, "tcPr");
  const generatedProperties = wordChildren(generatedCell, "tcPr");
  if (sourceProperties.length > 1 || generatedProperties.length > 1)
    return false;
  const sourceProperty = sourceProperties.at(0);
  const generatedProperty = generatedProperties.at(0);
  const sourceView = sourceXml.slice(sourceRange.start, sourceRange.end);
  let generatedView = serializedXml.slice(
    generatedRange.start,
    generatedRange.end,
  );
  const baselineContent = [...before];
  const currentContent: BlockContent[] = [...after];
  if (sourceProperty !== undefined) {
    if (
      getChildElements(sourceCell).at(0) !== sourceProperty ||
      !generatedProperty ||
      getChildElements(generatedCell).at(0) !== generatedProperty
    )
      return false;
    const range = getXmlSourceRange(sourceProperty);
    if (!range) return false;
    const propertyXml = materializeReplayFragment({
      xml: sourceXml.slice(range.start, range.end),
      element: sourceProperty,
      sourceNamespace: getNamespaceUri(sourceCell) ?? "",
    });
    if (propertyXml === null) return false;
    // Wrapper metadata is unchanged. Treat it as an explicitly owned opaque
    // child so source offsets, lexical syntax and surrounding gaps stay exact.
    const property = {
      type: "preservedBlock",
      xml: propertyXml,
    } as const satisfies BlockContent;
    baselineContent.unshift(property);
    currentContent.unshift(property);
  } else if (generatedProperty !== undefined) {
    const range = getXmlSourceRange(generatedProperty);
    if (!range) return false;
    const stripped = spliceXml(generatedView, [
      {
        start: range.start - generatedRange.start,
        end: range.end - generatedRange.start,
        newXml: "",
      },
    ]);
    if (stripped === null) return false;
    generatedView = stripped;
  }
  const patched = replayBlocks(
    {
      sourceXml: sourceView,
      serializedXml: generatedView,
      baselineContent,
      currentContent,
    },
    {
      source: sourceCell.namespaceScope,
      generated: generatedCell.namespaceScope,
      generatedRoot,
    },
  );
  if (patched === null) return false;
  splices.push({
    start: sourceRange.start - sourceRangeStart,
    end: sourceRange.end - sourceRangeStart,
    newXml: patched,
  });
  return true;
};

const readDocumentStory = (xml: string) => {
  const document = readStory(xml);
  if (!document || getLocalName(document.root.name) !== "document") return null;
  const bodies = wordChildren(document.root, "body");
  const body = bodies.at(0);
  if (
    bodies.length !== 1 ||
    !body ||
    wordChildren(body, "sectPr").length > 1 ||
    wordChildren(document.root, "background").length > 1
  )
    return null;
  const range = getXmlSourceRange(body);
  return range ? { document, body, range } : null;
};

/** The trusted source part as every save of it reads it. */
type SourceDocument = {
  readonly bodyRange: XmlRange;
  readonly bodyScope: XmlNamespaceScope | undefined;
  readonly story: SourceStory;
};

// A source part is read once per captured baseline (null: it cannot be read or
// does not match the baseline). The key is the private captured body, whose
// fingerprint getDocumentSourceBaseline verifies before a save passes it here,
// so neither the part nor its baseline can have changed since.
const sourceDocuments = new WeakMap<
  DocumentBody,
  { readonly xml: string; readonly read: SourceDocument | null }
>();

const readSourceDocument = (
  sourceXml: string,
  baseline: DocumentBody,
): SourceDocument | null => {
  const cached = sourceDocuments.get(baseline);
  if (cached?.xml === sourceXml) return cached.read;
  const source = readDocumentStory(sourceXml);
  const body = source && storyOfElement(source.body, "sectPr");
  const story = body && readSourceStory(body, baseline.content, "sectPr");
  const read =
    source && story
      ? {
          bodyRange: source.range,
          bodyScope: source.body.namespaceScope,
          story,
        }
      : null;
  sourceDocuments.set(baseline, { xml: sourceXml, read });
  return read;
};

/** Validate source/model correspondence before replaying an unchanged complete part. */
export const canReplayDocumentSource = (
  sourceXml: string,
  baseline: DocumentBody,
): boolean => readSourceDocument(sourceXml, baseline) !== null;

type DocumentBlockReplayOptions = {
  sourceXml: string;
  serializedXml: string;
  baseline: DocumentBody;
  current: DocumentBody;
};

/**
 * Replay body blocks and modeled document metadata inside the authored part
 * shell. The generated part is read once; the source part once per baseline.
 *
 * The answer carries the replayed body blocks when those are all it changed.
 * It carries none when the replay also rewrote the section properties, the
 * background or a self-closing body, which are not children of the body.
 */
export const buildDocumentBlockReplay = ({
  sourceXml,
  serializedXml,
  baseline,
  current,
}: DocumentBlockReplayOptions): SerializedPart | null => {
  const source = readSourceDocument(sourceXml, baseline);
  const generatedStory = readDocumentStory(serializedXml);
  if (!source || !generatedStory) return null;
  const { bodyRange: sourceRange, bodyScope: sourceBodyScope } = source;
  const { document: generated, body: generatedBody } = generatedStory;
  const generatedBlocks = storyOfElement(generatedBody, "sectPr");
  if (!generatedBlocks) return null;
  const bodySplices = replayStoryBlocks({
    sourceXml,
    serializedXml,
    source: source.story,
    generated: generatedBlocks,
    baselineContent: baseline.content,
    currentContent: current.content,
    generatedRoot: generated.root,
  });
  if (bodySplices === null) return null;
  // The splices that change the source; the others copy a block's own bytes.
  const replayed = bodySplices.filter(
    ({ start, end, newXml }) => newXml !== sourceXml.slice(start, end),
  );
  let blocksOnly = replayed.every(
    ({ start, end }) => start !== sourceRange.start || end !== sourceRange.end,
  );
  let bodyXml = spliceXml(
    sourceXml.slice(sourceRange.start, sourceRange.end),
    bodySplices.map(({ start, end, newXml }) => ({
      start: start - sourceRange.start,
      end: end - sourceRange.start,
      newXml,
    })),
  );
  if (bodyXml === null) return null;
  const replaceMetadata = (
    xml: string,
    sourceContainer: XmlElement,
    generatedContainer: XmlElement,
    name: string,
  ): string | null => {
    const before = wordChildren(sourceContainer, name);
    const after = wordChildren(generatedContainer, name);
    if (before.length > 1 || after.length > 1) return null;
    const old = before.at(0);
    const next = after.at(0);
    const containerRange = getXmlSourceRange(sourceContainer);
    const oldRange = old && getXmlSourceRange(old);
    const nextRange = next && getXmlSourceRange(next);
    if (!containerRange || (old && !oldRange) || (next && !nextRange))
      return null;
    const replacement =
      next && nextRange
        ? materializeReplayFragment({
            xml: serializedXml.slice(nextRange.start, nextRange.end),
            element: next,
            sourceNamespace: getNamespaceUri(sourceContainer) ?? "",
            generatedRoot: generated.root,
          })
        : "";
    if (replacement === null) return null;
    if (oldRange)
      return spliceXml(xml, [
        {
          start: oldRange.start - containerRange.start,
          end: oldRange.end - containerRange.start,
          newXml: replacement,
        },
      ]);
    if (!next) return xml;
    const insertion =
      name === "background"
        ? ([...xml.matchAll(XML_TOKEN)].at(0)?.[0].length ?? -1)
        : xml.lastIndexOf("</");
    if (insertion < 0) return null;
    return spliceXml(xml, [
      { start: insertion, end: insertion, newXml: replacement },
    ]);
  };
  if (
    canonicalJson(baseline.finalSectionProperties) !==
    canonicalJson(current.finalSectionProperties)
  ) {
    const reparsed = parseStreamingXmlWithSourceRanges(
      bodyXml,
      sourceBodyScope,
    );
    if (reparsed.status !== "parsed") return null;
    const root = getChildElements(reparsed.value).at(0);
    if (!root) return null;
    bodyXml = replaceMetadata(bodyXml, root, generatedBody, "sectPr");
    if (bodyXml === null) return null;
    blocksOnly = false;
  }
  let result = spliceXml(sourceXml, [
    { start: sourceRange.start, end: sourceRange.end, newXml: bodyXml },
  ]);
  if (result === null) return null;
  if (
    canonicalJson(baseline.background) !== canonicalJson(current.background)
  ) {
    const reparsed = readStory(result);
    if (!reparsed) return null;
    // Root range can begin after an XML declaration: keep those absolute offsets.
    const rootXml = result.slice(reparsed.range.start, reparsed.range.end);
    const replaced = replaceMetadata(
      rootXml,
      reparsed.root,
      generated.root,
      "background",
    );
    if (replaced === null) return null;
    result = spliceXml(result, [
      {
        start: reparsed.range.start,
        end: reparsed.range.end,
        newXml: replaced,
      },
    ]);
    if (result === null) return null;
    blocksOnly = false;
  }
  return {
    xml: result,
    replay: blocksOnly ? { source: sourceXml, splices: replayed } : undefined,
  };
};
