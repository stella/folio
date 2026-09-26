import { REVIEW_CARRIERS } from "@stll/docx-core/model";
import { Fragment, Slice, type Node as PMNode } from "prosemirror-model";
import { Mapping, StepMap, ReplaceStep, type Step } from "prosemirror-transform";
import {
  getProseParagraphPropertySourceToken,
  PROSE_PARAGRAPH_SOURCE_TOKEN_ATTR,
  recreateProseNodeWithParagraphPropertySource as rebuild,
} from "../docx/paragraphPropertySource";
import { expectParagraphAttrs } from "../prosemirror/attrs";
import { composeRevisionResolutionMaps } from "./revisionResolutionStep";
import { recordNodeResolution } from "./revisionResolutionEdits";
import { resolveAllTableChanges } from "../prosemirror/commands/resolveAllTableChanges";
import { resolveParagraphChangeAttrs } from "../prosemirror/commands/resolveParagraphProperties";
import { resolveAllNodePropertyChangeAttrs } from "../prosemirror/commands/resolveNodePropertyChangeAttrs";
import { holdsNoContent } from "../prosemirror/zeroWidthAnchors";
import { paragraphRunStyleContext, type RunStyleResolver } from "../prosemirror/runStyleFormatting";
import {
  readAuthoredRunFormatting,
  reconcileRunFormattingMarks,
} from "../prosemirror/runFormattingReconciliation";
import { expandRunFormattingCarrier } from "../prosemirror/runFormattingInlineCarriers";
import { resolveInlineRevisions, type RevisionResolutionMode } from "./revisionResolutionInline";
import type { RemovedSectionReference } from "./sectionEndpointResolution";

// Property changes do not change positions. Resolve their style context before
// visiting the runs whose inherited formatting depends on it.
type ResolvePropertiesOptions = {
  doc: PMNode;
  mode: RevisionResolutionMode;
  styleResolver: RunStyleResolver | null;
  steps: Step[];
  changedRanges: { from: number; to: number }[];
};
const resolveProperties = ({
  doc,
  mode,
  styleResolver,
  steps,
  changedRanges,
}: ResolvePropertiesOptions) => {
  let structural = false;
  const walk = (node: PMNode, position: number): PMNode => {
    if (node.isText) return node;
    const nextAttrs =
      node.type.name === "paragraph"
        ? resolveParagraphChangeAttrs({
            node,
            mode,
            boundaryCovered: true,
            revisionSet: null,
            styleResolver,
          })
        : null;
    let resolved = node;
    if (nextAttrs) resolved = rebuild(node, { attrs: nextAttrs });
    else if (node.type.name !== "paragraph")
      resolved = resolveAllNodePropertyChangeAttrs(node, mode);
    if (resolved !== node) {
      if (node.type.name === "paragraph")
        changedRanges.push({ from: position, to: position + node.nodeSize });
      structural ||= node.type.name !== "paragraph";
    }
    if (position >= 0) recordNodeResolution({ before: node, after: resolved, position, steps });
    const rebased = new Map<PMNode, PMNode>();
    if (
      node.type.name === "paragraph" &&
      styleResolver &&
      nextAttrs &&
      nextAttrs["styleId"] !== node.attrs["styleId"]
    ) {
      const previousContext = paragraphRunStyleContext(node, styleResolver);
      const nextContext = paragraphRunStyleContext(resolved, styleResolver);
      node.descendants((inline, pos) => {
        if (!inline.isInline) return true;
        const carrier = expandRunFormattingCarrier(inline, pos);
        if (!carrier) return !inline.isAtom;
        for (const { node: representation } of carrier.representations) {
          if (
            mode === "reject" &&
            representation.marks.some(({ type }) => type.name === "deletion") &&
            representation.marks.some(({ type }) => type.name === "runFormattingOverride")
          )
            continue;
          const authoredFormatting = readAuthoredRunFormatting({
            context: previousContext,
            marks: representation.marks,
            styleResolver,
          });
          const marks = reconcileRunFormattingMarks({
            authoredFormatting,
            context: nextContext,
            node: representation,
            styleResolver,
          });
          rebased.set(representation, representation.mark(marks));
        }
        return false;
      });
    }
    const children: PMNode[] = [];
    let changed = false;
    const visitChild = (child: PMNode, childPosition: number): PMNode => {
      const replacement = rebased.get(child) ?? child;
      if (rebased.size === 0) return walk(replacement, childPosition);
      recordNodeResolution({ before: child, after: replacement, position: childPosition, steps });
      if (replacement.isLeaf) return replacement;
      const nested: PMNode[] = [];
      replacement.forEach((inner, offset) =>
        nested.push(visitChild(inner, childPosition + 1 + offset)),
      );
      return rebuild(replacement, { content: nested });
    };
    resolved.forEach((child, offset) => {
      const result = visitChild(child, position + 1 + offset);
      children.push(result);
      changed ||= result !== child;
    });
    return changed ? rebuild(resolved, { content: children }) : resolved;
  };
  return { resolved: walk(doc, -1), structural };
};

type StructuralContext = {
  mode: RevisionResolutionMode;
  deleted: { from: number; to: number }[];
  steps: Step[];
  changedRanges: { from: number; to: number }[];
  replacements: { from: number; to: number; slice: Slice }[];
  transfers: { displacedToken: string | null; selectedToken: string | null }[];
  tableRanges: {
    node: PMNode;
    inputNode: PMNode;
    sourceNodeSize: number;
    position: number;
    inputMap: StepMap;
    positionMap: StepMap;
    ranges: readonly { from: number; to: number }[];
    paragraphOffsets: readonly { source: number; final: number }[];
  }[];
  structural: boolean;
  failed: boolean;
  removedEndpointCount: number;
  removedReferences: RemovedSectionReference[];
};

type StructuralReplacement = { from: number; to: number; slice: Slice };

/** Use a table's granular map in place of its enclosing replay replacement. */
type StructuralReplacementMapOptions = {
  replacements: readonly StructuralReplacement[];
  tables: readonly StructuralContext["tableRanges"][number][];
  origin: number;
};
const structuralReplacementMap = ({
  replacements,
  tables,
  origin,
}: StructuralReplacementMapOptions): StepMap => {
  const tablesByPosition = new Map(tables.map((table) => [table.position, table]));
  const ranges: number[] = [];
  const mapFor = ({ from, to, slice }: StructuralReplacement): StepMap => {
    const table = tablesByPosition.get(from);
    if (table && table.sourceNodeSize === to - from) {
      const tableRanges: number[] = [];
      table.positionMap.forEach((oldStart, oldEnd, newStart, newEnd) => {
        tableRanges.push(from - origin + oldStart, oldEnd - oldStart, newEnd - newStart);
      });
      return new StepMap(tableRanges);
    }
    return new StepMap([from - origin, to - from, slice.size]);
  };
  const isDeletion = ({ from, to, slice }: StructuralReplacement): boolean =>
    slice.size === 0 && !tablesByPosition.has(from) && to > from;
  const groupMap = (group: readonly StructuralReplacement[]): StepMap => {
    if (group.every(isDeletion)) {
      const first = group.at(0);
      if (!first) return StepMap.empty;
      let end = first.to;
      for (const replacement of group.slice(1)) {
        end += replacement.to - replacement.from;
      }
      return new StepMap([first.from - origin, end - first.from, 0]);
    }
    let maps = group.toReversed().map(mapFor);
    while (maps.length > 1) {
      const combined: StepMap[] = [];
      for (let index = 0; index < maps.length; index += 2) {
        const first = maps.at(index);
        const second = maps.at(index + 1);
        if (!first) break;
        combined.push(second ? composeRevisionResolutionMaps(first, second) : first);
      }
      maps = combined;
    }
    return maps.at(0) ?? StepMap.empty;
  };
  let group: StructuralReplacement[] = [];
  let influenceEnd = -1;
  const flush = (): void => {
    const map = groupMap(group);
    map.forEach((oldStart, oldEnd, newStart, newEnd) => {
      ranges.push(oldStart, oldEnd - oldStart, newEnd - newStart);
    });
  };
  for (const replacement of replacements) {
    if (group.length > 0 && replacement.from >= influenceEnd) {
      flush();
      group = [];
    }
    if (group.length === 0) {
      group.push(replacement);
      influenceEnd = replacement.to;
      continue;
    }
    group.push(replacement);
    // The old span bounds how far a higher-address replacement can extend
    // the source region consumed by an earlier overlapping operation.
    influenceEnd += replacement.to - replacement.from;
  }
  if (group.length > 0) {
    flush();
  }
  return new StepMap(ranges);
};
const removedEndpoint = (node: PMNode, context: StructuralContext): void => {
  const properties = expectParagraphAttrs(node)._sectionProperties;
  if (properties === undefined) return;
  context.removedEndpointCount++;
  for (const { type, rId } of properties?.headerReferences ?? [])
    context.removedReferences.push({ part: "header", type, relationshipId: rId });
  for (const { type, rId } of properties?.footerReferences ?? [])
    context.removedReferences.push({ part: "footer", type, relationshipId: rId });
};

type ParagraphChain = { node: PMNode; chunks: Fragment[]; position: number; before: PMNode };
type ResolveStructureOptions = { node: PMNode; position: number; context: StructuralContext };
const resolveStructure = ({ node, position, context }: ResolveStructureOptions): PMNode | null => {
  if (node.isLeaf) return node;
  const structuralStart = {
    steps: context.steps.length,
    deleted: context.deleted.length,
    replacements: context.replacements.length,
    tableRanges: context.tableRanges.length,
  };
  const entries: { node: PMNode | null; original: PMNode; position: number }[] = [];
  node.forEach((child, offset) =>
    entries.push({
      node: resolveStructure({ node: child, position: position + 1 + offset, context }),
      original: child,
      position: position + 1 + offset,
    }),
  );
  const reversed: PMNode[] = [];
  let chain: ParagraphChain | null = null;
  let followingContainerChild = false;
  const flush = (): void => {
    if (!chain) return;
    const children: PMNode[] = [];
    for (let index = chain.chunks.length - 1; index >= 0; index--)
      chain.chunks[index]?.forEach((child) => children.push(child));
    const final = rebuild(chain.node, { content: Fragment.fromArray(children) });
    if (!chain.before.sameMarkup(final))
      context.changedRanges.push({
        from: chain.position,
        to: chain.position + chain.before.nodeSize,
      });
    recordNodeResolution({
      before: chain.before,
      after: final,
      position: chain.position,
      steps: context.steps,
    });
    reversed.push(chain.chunks.length === 1 && final.eq(chain.before) ? chain.before : final);
    chain = null;
  };
  for (let index = entries.length - 1; index >= 0; index--) {
    const entry = entries[index];
    if (!entry) continue;
    let paragraph = entry.node;
    if (paragraph === null) {
      // Break the loop's inference cycle through filler, paragraph, and chain.
      const hasSibling: boolean = index + reversed.length + (chain ? 1 : 0) > 0;
      const filler: PMNode | null | undefined = hasSibling
        ? null
        : node.type.schema.nodes["paragraph"]?.createAndFill();
      context.replacements.push({
        from: entry.position,
        to: entry.position + entry.original.nodeSize,
        slice: filler ? new Slice(Fragment.from(filler), 0, 0) : Slice.empty,
      });
      if (!filler) continue;
      paragraph = filler;
    }
    if (paragraph.type.name !== "paragraph") {
      flush();
      reversed.push(paragraph);
      followingContainerChild ||= paragraph.type.spec["tableRole"] === "table";
      continue;
    }
    const marker = expectParagraphAttrs(paragraph).pPrMark;
    const markWasAdded = marker?.kind === "ins" || marker?.kind === "moveTo";
    const joins = marker && markWasAdded !== (context.mode === "accept");
    if (joins && chain) {
      const empty = holdsNoContent(paragraph);
      const owner = empty ? chain.node : paragraph;
      const next = chain.node;
      if (empty) {
        const displacedToken = getProseParagraphPropertySourceToken(paragraph);
        const selectedToken = getProseParagraphPropertySourceToken(next);
        context.transfers.push({
          displacedToken: typeof displacedToken === "string" ? displacedToken : null,
          selectedToken: typeof selectedToken === "string" ? selectedToken : null,
        });
      }
      chain.node = rebuild(owner, {
        attrs: {
          ...owner.attrs,
          pPrMark: next.attrs["pPrMark"],
          _sectionProperties: next.attrs["_sectionProperties"],
          ...(empty
            ? {
                [PROSE_PARAGRAPH_SOURCE_TOKEN_ATTR]:
                  getProseParagraphPropertySourceToken(next) ?? null,
              }
            : {}),
        },
      });
      chain.chunks.push(paragraph.content);
      chain.position = entry.position;
      chain.before = paragraph;
      context.deleted.push({
        from: entry.position + paragraph.nodeSize - 1,
        to: entry.position + paragraph.nodeSize + 1,
      });
      removedEndpoint(paragraph, context);
      continue;
    }
    flush();
    if (
      joins &&
      holdsNoContent(paragraph) &&
      (markWasAdded || followingContainerChild) &&
      index + reversed.length + 1 > 1
    ) {
      context.deleted.push({ from: entry.position, to: entry.position + paragraph.nodeSize });
      removedEndpoint(paragraph, context);
      continue;
    }
    const resolved: PMNode = marker
      ? rebuild(paragraph, { attrs: { ...paragraph.attrs, pPrMark: null } })
      : paragraph;
    chain = {
      node: resolved,
      chunks: [resolved.content],
      position: entry.position,
      before: paragraph,
    };
    followingContainerChild = true;
  }
  flush();
  reversed.reverse();
  if (node.type.name === "doc") {
    const carrier = reversed.at(-1);
    if (
      carrier?.type.name === "paragraph" &&
      carrier.attrs["reviewCarrier"] === REVIEW_CARRIERS.TERMINAL_TABLE
    ) {
      const table = reversed.findLast((child) => child.type.name === "table");
      let pending = false;
      table?.descendants((child) => {
        pending ||=
          (child.type.name === "tableRow" &&
            child.attrs["trDel"] !== null &&
            typeof child.attrs["trDel"] === "object") ||
          (child.type.name !== "tableRow" &&
            child.marks.some((mark) => mark.type.name === "deletion"));
        return !pending;
      });
      const entry = entries.at(-1);
      if (!pending && entry) {
        if (context.mode === "accept") {
          const cleared = rebuild(carrier, {
            attrs: { ...carrier.attrs, reviewCarrier: undefined },
          });
          recordNodeResolution({
            before: carrier,
            after: cleared,
            position: entry.position,
            steps: context.steps,
          });
          context.changedRanges.push({
            from: entry.position,
            to: entry.position + carrier.nodeSize,
          });
          reversed[reversed.length - 1] = cleared;
        } else if (reversed.length > 1) {
          reversed.pop();
          context.deleted.push({ from: entry.position, to: entry.position + carrier.nodeSize });
        } else {
          const filler = node.type.schema.nodes["paragraph"]?.createAndFill();
          if (filler) {
            reversed[0] = filler;
            context.replacements.push({
              from: entry.position,
              to: entry.position + carrier.nodeSize,
              slice: new Slice(Fragment.from(filler), 0, 0),
            });
          }
        }
      }
    }
  }
  const rebuilt = rebuild(node, { content: reversed });
  if (node.type.spec["tableRole"] !== "table") return rebuilt;
  const tableResult = resolveAllTableChanges({ table: rebuilt, mode: context.mode });
  context.structural ||= tableResult.structural;
  context.failed ||= tableResult.failed;
  const final = tableResult.node;
  if (final) {
    const inside = [
      ...context.replacements.slice(structuralStart.replacements),
      ...context.deleted
        .slice(structuralStart.deleted)
        .map(({ from, to }) => ({ from, to, slice: Slice.empty })),
    ].sort((a, b) => a.from - b.from);
    const inputMap = structuralReplacementMap({
      replacements: inside,
      tables: context.tableRanges.slice(structuralStart.tableRanges),
      origin: position,
    });
    const positionMap = composeRevisionResolutionMaps(inputMap, tableResult.positionMap);
    context.tableRanges.push({
      node: final,
      inputNode: rebuilt,
      sourceNodeSize: node.nodeSize,
      position,
      inputMap,
      positionMap,
      ranges: tableResult.changedParagraphRanges,
      paragraphOffsets: tableResult.paragraphOffsets,
    });
  }
  context.steps.length = structuralStart.steps;
  context.deleted.length = structuralStart.deleted;
  context.replacements.length = structuralStart.replacements;
  if (final && !node.eq(final))
    context.replacements.push({
      from: position,
      to: position + node.nodeSize,
      slice: new Slice(Fragment.from(final), 0, 0),
    });
  return final;
};

/** Resolve a story in linear passes, without rebuilding its root per revision. */
type ResolveWholeStoryOptions = {
  doc: PMNode;
  mode: RevisionResolutionMode;
  styleResolver: RunStyleResolver | null;
};
export const resolveWholeStory = ({ doc, mode, styleResolver }: ResolveWholeStoryOptions) => {
  const steps: Step[] = [];
  const propertyRanges: { from: number; to: number }[] = [];
  const propertyResult = resolveProperties({
    doc,
    mode,
    styleResolver,
    steps,
    changedRanges: propertyRanges,
  });
  const properties = propertyResult.resolved;
  const inline = resolveInlineRevisions({
    doc: properties,
    mode,
    styleResolver,
    keepType: doc.type.schema.marks[mode === "accept" ? "insertion" : "deletion"],
    removeType: doc.type.schema.marks[mode === "accept" ? "deletion" : "insertion"],
  });
  const context: StructuralContext = {
    mode,
    steps: [],
    changedRanges: [],
    deleted: [],
    replacements: [],
    transfers: [],
    tableRanges: [],
    structural: propertyResult.structural,
    failed: false,
    removedEndpointCount: 0,
    removedReferences: [],
  };
  const resolved =
    resolveStructure({ node: inline?.resolved ?? properties, position: -1, context }) ?? doc;
  if (inline) steps.push(...inline.steps);
  steps.push(...context.steps);
  const replacements = [
    ...context.replacements,
    ...context.deleted.map(({ from, to }) => ({ from, to, slice: Slice.empty })),
  ].sort((a, b) => a.from - b.from);
  const structuralMap = structuralReplacementMap({
    replacements,
    tables: context.tableRanges,
    origin: 0,
  });
  for (const { from, to, slice } of replacements.toReversed())
    steps.push(new ReplaceStep(from, to, slice));
  const positionMap = composeRevisionResolutionMaps(
    inline?.positionMap ?? StepMap.empty,
    structuralMap,
  );
  const mapping = new Mapping([positionMap]);
  const changedRanges = [...propertyRanges, ...(inline?.ranges ?? [])];
  return {
    resolved,
    mapping,
    positionMap,
    steps,
    changedRanges,
    inlineMap: inline?.positionMap ?? StepMap.empty,
    tableRanges: context.tableRanges,
    transfers: context.transfers,
    structuralRanges: context.changedRanges,
    structuralMap,
    structural: context.structural,
    failed: context.failed,
    removedEndpointCount: context.removedEndpointCount,
    removedReferences: context.removedReferences,
  };
};
