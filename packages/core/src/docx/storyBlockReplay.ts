import { escapeXmlAttribute } from "@stll/docx-core";
import { DOCX_CONFORMANCE_CLASSES } from "@stll/docx-core/model";
import type { BlockContent } from "../types/document";
import { captureSourceProfileXml, isSafeCapturedXmlDocument } from "./verbatimCapture";
import { canonicalJson } from "../utils/canonicalJson";
import { spliceXml, type XmlSplice } from "./selectiveXmlPatch";
import { TRANSITIONAL_NAMESPACE_BY_STRICT_URI } from "./strictValueEncodings.gen";
import { getXmlSourceRange, parseStreamingXmlWithSourceRanges } from "./streamingXmlParser";
import {
  getAttributeByNamespaceUri,
  getChildElements,
  getLocalName,
  getNamespaceUri,
  WORDPROCESSINGML_NAMESPACE_URIS,
  type XmlElement,
  type XmlNamespaceScope,
} from "./xmlParser";

const BLOCK_ELEMENT_NAMES = {
  paragraph: "p",
  table: "tbl",
  blockSdt: "sdt",
  blockCustomXml: "customXml",
  preservedBlock: null,
  bookmarkStart: "bookmarkStart",
  bookmarkEnd: "bookmarkEnd",
} as const satisfies Record<BlockContent["type"], string | null>;
type ModelMatchOptions = {
  value: BlockContent | undefined;
  element: XmlElement;
  mode: "source" | "generated";
};
const modelMatchesElement = ({ value, element, mode }: ModelMatchOptions): boolean => {
  if (value === undefined) return false;
  // The parser projects a supported AlternateContent branch as one block;
  // source cardinality and the trusted baseline preserve its outer carrier.
  if (
    mode === "source" &&
    getLocalName(element.name) === "AlternateContent" &&
    getNamespaceUri(element) === "http://schemas.openxmlformats.org/markup-compatibility/2006" &&
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
  const parsed = parseStreamingXmlWithSourceRanges(value.xml, element.namespaceScope);
  if (parsed.status !== "parsed") return false;
  const roots = getChildElements(parsed.value);
  const root = roots.at(0);
  const normalize = (uri: string | undefined) =>
    uri === undefined ? undefined : (TRANSITIONAL_NAMESPACE_BY_STRICT_URI.get(uri) ?? uri);
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

const readStory = (xml: string, scope?: XmlNamespaceScope) => {
  const parsed = parseStreamingXmlWithSourceRanges(xml, scope);
  if (parsed.status !== "parsed") return null;
  const roots = getChildElements(parsed.value);
  const root = roots.at(0);
  if (
    roots.length !== 1 ||
    !root ||
    !WORDPROCESSINGML_NAMESPACE_URIS.has(getNamespaceUri(root) ?? "")
  )
    return null;
  if (
    root.elements?.some((child) => child.type === "text" && String(child.text ?? "").trim() !== "")
  )
    return null;
  const children = getChildElements(root);
  const blocks = [];
  for (const element of children) {
    const range = getXmlSourceRange(element);
    if (!range) return null;
    blocks.push({ element, ...range });
  }
  const range = getXmlSourceRange(root);
  return range ? { root, blocks, range } : null;
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

type GeneratedFragmentOptions = {
  xml: string;
  element: XmlElement;
  sourceNamespace: string;
  generatedRoot?: XmlElement;
};

const generatedFragment = ({
  xml,
  element,
  sourceNamespace,
  generatedRoot,
}: GeneratedFragmentOptions): string | null => {
  const strict = TRANSITIONAL_NAMESPACE_BY_STRICT_URI.has(sourceNamespace);
  const tokens = [...xml.matchAll(XML_TOKEN)];
  const first = tokens.at(0);
  if (!first || first.index !== 0 || first[0].startsWith("<!") || first[0].startsWith("<?"))
    return null;
  const opening = first[0];
  const additions = [];
  const bindings = effectiveBindings(element);
  const markupCompatibilityUri = "http://schemas.openxmlformats.org/markup-compatibility/2006";
  const ignorableUris = new Set([markupCompatibilityUri]);
  const inheritedIgnorable =
    generatedRoot === undefined
      ? null
      : getAttributeByNamespaceUri(generatedRoot, ignorableUris, "Ignorable");
  const localIgnorable = getAttributeByNamespaceUri(element, ignorableUris, "Ignorable");
  const ignorable = [
    ...new Set(
      `${inheritedIgnorable ?? ""} ${localIgnorable ?? ""}`.trim().split(/\s+/u).filter(Boolean),
    ),
  ];
  let ignorableAttribute: string | undefined;
  if (ignorable.length > 0) {
    const localAttribute = Object.keys(element.attributes ?? {}).find((name) => {
      const colon = name.indexOf(":");
      return (
        colon > 0 &&
        getLocalName(name) === "Ignorable" &&
        bindings.get(name.slice(0, colon)) === markupCompatibilityUri
      );
    });
    const compatibilityPrefix =
      localAttribute?.slice(0, localAttribute.indexOf(":")) ??
      [...bindings].find(([prefix, uri]) => prefix !== "" && uri === markupCompatibilityUri)?.[0];
    if (compatibilityPrefix === undefined || ignorable.some((prefix) => !bindings.has(prefix)))
      return null;
    ignorableAttribute = `${compatibilityPrefix}:Ignorable`;
    if (localIgnorable === null)
      additions.push(` ${ignorableAttribute}="${escapeXmlAttribute(ignorable.join(" "))}"`);
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
  const close = opening.endsWith("/>") ? replacedOpening.length - 2 : replacedOpening.length - 1;
  replacedOpening =
    replacedOpening.slice(0, close) + additions.join("") + replacedOpening.slice(close);
  const materialized = replacedOpening + xml.slice(opening.length);
  if (!isSafeCapturedXmlDocument(materialized)) return null;
  if (!strict) return materialized;
  const parsed = parseStreamingXmlWithSourceRanges(materialized);
  if (parsed.status !== "parsed") return null;
  const root = getChildElements(parsed.value).at(0);
  return root === undefined ? null : captureSourceProfileXml(root, DOCX_CONFORMANCE_CLASSES.STRICT);
};

const paragraphIdentity = (value: BlockContent): string | undefined =>
  value.type === "paragraph" ? value.paraId?.toUpperCase() : undefined;

type CandidateQueue = { indices: number[]; cursor: number };
type AddCandidateOptions = { map: Map<string, CandidateQueue>; key: string; index: number };
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
};

/** Replace only mapped block ranges, retaining authored root syntax and every intervening gap. */
const replayBlocks = (
  { sourceXml, baselineContent, currentContent, serializedXml }: StoryBlockReplayOptions,
  scope?: ReplayScopeOptions,
): string | null => {
  const source = readStory(sourceXml, scope?.source);
  const generated = readStory(serializedXml, scope?.generated);
  if (
    !source ||
    !generated ||
    getLocalName(source.root.name) !== getLocalName(generated.root.name) ||
    source.blocks.length !== baselineContent.length ||
    generated.blocks.length !== currentContent.length
  )
    return null;
  if (
    source.blocks.some(
      (block, index) =>
        !modelMatchesElement({
          value: baselineContent[index],
          element: block.element,
          mode: "source",
        }),
    ) ||
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
  const generatedRoot = scope?.generatedRoot ?? generated.root;
  const fingerprints = baselineContent.map(canonicalJson);
  const identities = new Map<string, CandidateQueue>();
  const byFingerprint = new Map<string, CandidateQueue>();
  for (const [index, block] of baselineContent.entries()) {
    const fingerprint = fingerprints[index];
    if (fingerprint === undefined) return null;
    addCandidate({ map: byFingerprint, key: fingerprint, index });
    const identity = paragraphIdentity(block);
    if (identity !== undefined) addCandidate({ map: identities, key: identity, index });
  }
  const used = new Set<number>();
  const fragments: string[] = [];
  for (const [index, block] of currentContent.entries()) {
    const fingerprint = canonicalJson(block);
    const identity = paragraphIdentity(block);
    let originalIndex =
      takeCandidate(identity === undefined ? undefined : identities.get(identity), used) ??
      takeCandidate(byFingerprint.get(fingerprint), used);
    // Containers have no durable block id. An unchanged wrapper at the same
    // ordinal can still retain its source while its nested blocks are edited.
    if (
      originalIndex === undefined &&
      (block.type === "table" || block.type === "blockSdt" || block.type === "blockCustomXml") &&
      !used.has(index)
    ) {
      const baseline = baselineContent[index];
      if (baseline?.type === block.type) originalIndex = index;
    }
    if (originalIndex !== undefined) used.add(originalIndex);
    const original = originalIndex === undefined ? undefined : source.blocks[originalIndex];
    if (originalIndex !== undefined && original && fingerprints[originalIndex] === fingerprint) {
      fragments.push(sourceXml.slice(original.start, original.end));
      continue;
    }
    const replacement = generated.blocks[index];
    if (!replacement) return null;
    if (original && originalIndex !== undefined) {
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
    const fragment = generatedFragment({
      xml: serializedXml.slice(replacement.start, replacement.end),
      element: replacement.element,
      sourceNamespace: getNamespaceUri(source.root) ?? "",
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
      ...sourceXml.slice(source.range.start, source.range.end).matchAll(XML_TOKEN),
    ].at(0)?.[0];
    if (!opening || !source.root.name) return null;
    if (opening.endsWith("/>")) {
      splices.push({
        start: source.range.start,
        end: source.range.end,
        newXml: `${opening.slice(0, -2)}>${fragments.join("")}</${source.root.name}>`,
      });
    } else {
      const close = sourceXml.lastIndexOf("</", source.range.end);
      if (close < source.range.start) return null;
      splices.push({ start: close, end: close, newXml: fragments.join("") });
    }
  }
  return spliceXml(sourceXml, splices);
};

export const buildStoryBlockReplay = (options: StoryBlockReplayOptions): string | null =>
  replayBlocks(options);

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
        sourceXml: sourceXml.slice(sourceContentRange.start, sourceContentRange.end),
        serializedXml: serializedXml.slice(generatedContentRange.start, generatedContentRange.end),
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
        if (!nestedContent(sourceElement, generatedElement, baseline.content, current.content))
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
        !nestedContent(sourceContainer, generatedContainer, baseline.content, current.content)
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
      if (sourceRows.length !== current.rows.length || generatedRows.length !== current.rows.length)
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
        if (sourceCells.length !== row.cells.length || generatedCells.length !== row.cells.length)
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
      return null;
    default: {
      const unreachable: never = current;
      return unreachable;
    }
  }
  return spliceXml(sourceXml.slice(sourceRange.start, sourceRange.end), splices);
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
  if (sourceProperties.length > 1 || generatedProperties.length > 1) return false;
  const sourceProperty = sourceProperties.at(0);
  const generatedProperty = generatedProperties.at(0);
  const sourceView = sourceXml.slice(sourceRange.start, sourceRange.end);
  let generatedView = serializedXml.slice(generatedRange.start, generatedRange.end);
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
    const propertyXml = generatedFragment({
      xml: sourceXml.slice(range.start, range.end),
      element: sourceProperty,
      sourceNamespace: getNamespaceUri(sourceCell) ?? "",
    });
    if (propertyXml === null) return false;
    // Wrapper metadata is unchanged. Treat it as an explicitly owned opaque
    // child so source offsets, lexical syntax and surrounding gaps stay exact.
    const property = { type: "preservedBlock", xml: propertyXml } as const satisfies BlockContent;
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
    { sourceXml: sourceView, serializedXml: generatedView, baselineContent, currentContent },
    { source: sourceCell.namespaceScope, generated: generatedCell.namespaceScope, generatedRoot },
  );
  if (patched === null) return false;
  splices.push({
    start: sourceRange.start - sourceRangeStart,
    end: sourceRange.end - sourceRangeStart,
    newXml: patched,
  });
  return true;
};
