import { Fragment, type Node as PMNode, Slice } from "prosemirror-model";
import {
  Mapping,
  ReplaceStep,
  type Step,
  StepMap,
  Transform,
} from "prosemirror-transform";

import { REVIEW_CARRIERS } from "@stll/docx-core/model";
import { paragraphMarkOutcome } from "@stll/docx-core/ops";

import type { NumberingMap } from "../docx/numberingParser";
import {
  getProseParagraphPropertySourceToken,
  recreateProseNodeWithParagraphPropertySource as rebuild,
} from "../docx/paragraphPropertySource";
import {
  anchoredTextBoxId,
  droppedTextBoxAnchorIds,
} from "../prosemirror/anchoredTextBoxes";
import { expectParagraphAttrs } from "../prosemirror/attrs";
import { inlineBookmarksOf } from "../prosemirror/commands/paragraphBookmarkJoin";
import { resolveAllTableChanges } from "../prosemirror/commands/resolveAllTableChanges";
import { resolveAllNodePropertyChangeAttrs } from "../prosemirror/commands/resolveNodePropertyChangeAttrs";
import { resolveParagraphChangeAttrs } from "../prosemirror/commands/resolveParagraphProperties";
import { resolveCommentReferences } from "../prosemirror/commentReferenceIntegrity";
import {
  isBlockBoundaryNode,
  paragraphRemainsOf,
} from "../prosemirror/keptRangeBoundaries";
import { rebaseParagraphRunContent } from "../prosemirror/rebaseParagraphRuns";
import { createResolvedCommentPointCollector } from "../prosemirror/replacedAnnotations";
import { expandRunFormattingCarrier } from "../prosemirror/runFormattingInlineCarriers";
import {
  readAuthoredRunFormatting,
  reconcileRunFormattingMarks,
} from "../prosemirror/runFormattingReconciliation";
import {
  paragraphRunStyleContext,
  type RunStyleResolver,
} from "../prosemirror/runStyleFormatting";
import {
  runParagraphsIntoTables,
  runsIntoFollowingTable,
} from "../prosemirror/tableRunIn";
import { paragraphContentOf } from "../prosemirror/zeroWidthAnchors";
import { recordNodeResolution } from "./revisionResolutionEdits";
import {
  resolveInlineRevisions,
  type RevisionResolutionMode,
} from "./revisionResolutionInline";
import { composeRevisionResolutionMaps } from "./revisionResolutionStep";
import type { RemovedSectionReference } from "./sectionEndpointResolution";

// Property changes do not change positions. Resolve their style context before
// visiting the runs whose inherited formatting depends on it.
type ResolvePropertiesOptions = {
  doc: PMNode;
  mode: RevisionResolutionMode;
  styleResolver: RunStyleResolver | null;
  numbering: NumberingMap | null;
  steps: Step[];
  changedRanges: { from: number; to: number }[];
};
const resolveProperties = ({
  doc,
  mode,
  styleResolver,
  numbering,
  steps,
  changedRanges,
}: ResolvePropertiesOptions) => {
  const removedMark =
    doc.type.schema.marks[mode === "accept" ? "deletion" : "insertion"];
  const collector = createResolvedCommentPointCollector({
    removes: (node) => removedMark?.isInSet(node.marks) !== undefined,
    retiresStructure: () => mode === "accept",
  });
  let structural = false;
  let removedEndpointCount = 0;
  const removedReferences: RemovedSectionReference[] = [];
  // Noted on this walk so a story without one needs no separate scan.
  let runsIntoTable = false;
  const walk = ({
    node,
    position,
    eligibleComments = true,
    previous,
  }: {
    node: PMNode;
    position: number;
    eligibleComments?: boolean;
    previous?: PMNode;
  }): PMNode => {
    const childComments =
      eligibleComments && collector.visit(node, position, previous);
    if (node.isText) return node;
    if (!runsIntoTable && !node.isTextblock) {
      for (let index = 0; index + 1 < node.childCount; index++) {
        if (
          runsIntoFollowingTable({
            parent: node,
            index,
            mode,
          })
        ) {
          runsIntoTable = true;
          break;
        }
      }
    }
    const nextAttrs =
      node.type.name === "paragraph"
        ? resolveParagraphChangeAttrs({
            node,
            mode,
            boundaryCovered: true,
            revisionSet: null,
            styleResolver,
            numbering,
          })
        : null;
    let resolved = node;
    if (nextAttrs) {
      if (
        node.type.name === "paragraph" &&
        ownsSection(node) &&
        nextAttrs["_sectionProperties"] == null
      ) {
        const properties = expectParagraphAttrs(node)._sectionProperties;
        removedEndpointCount++;
        for (const { type, rId } of properties?.headerReferences ?? [])
          removedReferences.push({ part: "header", type, relationshipId: rId });
        for (const { type, rId } of properties?.footerReferences ?? [])
          removedReferences.push({ part: "footer", type, relationshipId: rId });
      }
      resolved = rebuild(node, { attrs: nextAttrs });
    } else if (node.type.name !== "paragraph")
      resolved = resolveAllNodePropertyChangeAttrs(node, mode);
    if (resolved !== node) {
      if (node.type.name === "paragraph")
        changedRanges.push({ from: position, to: position + node.nodeSize });
      structural ||= node.type.name !== "paragraph";
    }
    if (position >= 0)
      recordNodeResolution({ before: node, after: resolved, position, steps });
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
    const visitChild = (
      child: PMNode,
      childPosition: number,
      previousChild?: PMNode,
    ): PMNode => {
      const replacement = rebased.get(child) ?? child;
      if (rebased.size === 0)
        return walk({
          node: replacement,
          position: childPosition,
          eligibleComments: childComments,
          ...(previousChild ? { previous: previousChild } : {}),
        });
      if (childComments) collector.visit(child, childPosition, previousChild);
      recordNodeResolution({
        before: child,
        after: replacement,
        position: childPosition,
        steps,
      });
      if (replacement.isLeaf) return replacement;
      const nested: PMNode[] = [];
      replacement.forEach((inner, offset) =>
        nested.push(visitChild(inner, childPosition + 1 + offset)),
      );
      return rebuild(replacement, { content: nested });
    };
    resolved.forEach((child, offset, index) => {
      const result = visitChild(
        child,
        position + 1 + offset,
        index > 0 ? (node.maybeChild(index - 1) ?? undefined) : undefined,
      );
      children.push(result);
      changed ||= result !== child;
    });
    return changed ? rebuild(resolved, { content: children }) : resolved;
  };
  const resolved = walk({ node: doc, position: -1 });
  return {
    resolved,
    structural,
    runsIntoTable,
    removedEndpointCount,
    removedReferences,
    pointComments:
      mode === "accept"
        ? new Set(collector.finish().keys())
        : new Set<number>(),
  };
};

type StructuralContext = {
  mode: RevisionResolutionMode;
  styleResolver: RunStyleResolver | null;
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
  bookmarkJoins: { from: number; to: number; positionMap: StepMap }[];
  structural: boolean;
  failed: boolean;
  removedEndpointCount: number;
  removedReferences: RemovedSectionReference[];
  /** Text box anchors the inline pass removed: their boxes go with them. */
  droppedTextBoxAnchors: ReadonlySet<string>;
};

type StructuralReplacement = { from: number; to: number; slice: Slice };

/** Use a table's granular map in place of its enclosing replay replacement. */
type StructuralReplacementMapOptions = {
  replacements: readonly StructuralReplacement[];
  tables: readonly StructuralContext["tableRanges"][number][];
  bookmarkJoins: readonly StructuralContext["bookmarkJoins"][number][];
  origin: number;
};
const structuralReplacementMap = ({
  replacements,
  tables,
  bookmarkJoins,
  origin,
}: StructuralReplacementMapOptions): StepMap => {
  const tablesByPosition = new Map(
    tables.map((table) => [table.position, table]),
  );
  const bookmarkJoinsByPosition = new Map(
    bookmarkJoins.map((join) => [join.from, join]),
  );
  const ranges: number[] = [];
  const mapFor = ({ from, to, slice }: StructuralReplacement): StepMap => {
    const table = tablesByPosition.get(from);
    if (table && table.sourceNodeSize === to - from) {
      const tableRanges: number[] = [];
      table.positionMap.forEach((oldStart, oldEnd, newStart, newEnd) => {
        tableRanges.push(
          from - origin + oldStart,
          oldEnd - oldStart,
          newEnd - newStart,
        );
      });
      return new StepMap(tableRanges);
    }
    const bookmarkJoin = bookmarkJoinsByPosition.get(from);
    if (bookmarkJoin?.to === to) {
      const joinRanges: number[] = [];
      bookmarkJoin.positionMap.forEach((oldStart, oldEnd, newStart, newEnd) => {
        joinRanges.push(
          from - origin + oldStart,
          oldEnd - oldStart,
          newEnd - newStart,
        );
      });
      return new StepMap(joinRanges);
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
        combined.push(
          second ? composeRevisionResolutionMaps(first, second) : first,
        );
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
      // An adjacent removal and replacement belong to one change. A StepMap
      // reads their shared boundary against the first range only and otherwise
      // misses the change on the other side.
      const last = ranges.length - 3;
      const previousStart = ranges.at(last);
      const previousOldSize = ranges.at(last + 1);
      const previousNewSize = ranges.at(last + 2);
      const newSize = newEnd - newStart;
      if (
        last >= 0 &&
        previousStart !== undefined &&
        previousOldSize !== undefined &&
        previousNewSize !== undefined &&
        (previousNewSize === 0 || newSize === 0) &&
        previousStart + previousOldSize === oldStart
      ) {
        ranges[last + 1] = previousOldSize + oldEnd - oldStart;
        ranges[last + 2] = previousNewSize + newSize;
        return;
      }
      ranges.push(oldStart, oldEnd - oldStart, newSize);
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
    context.removedReferences.push({
      part: "header",
      type,
      relationshipId: rId,
    });
  for (const { type, rId } of properties?.footerReferences ?? [])
    context.removedReferences.push({
      part: "footer",
      type,
      relationshipId: rId,
    });
};

/** A cleared record (null) and an absent one (undefined) both end no section. */
const ownsSection = (paragraph: PMNode): boolean =>
  expectParagraphAttrs(paragraph)._sectionProperties != null;

const precedingSectionOwner = (
  preceding: PMNode | null | undefined,
  properties: ReturnType<typeof expectParagraphAttrs>["_sectionProperties"],
): "unavailable" | "available" | "same" | "occupied" => {
  if (preceding?.type.name !== "paragraph") return "unavailable";
  if (!ownsSection(preceding)) return "available";
  const precedingProperties =
    expectParagraphAttrs(preceding)._sectionProperties;
  return JSON.stringify(precedingProperties) === JSON.stringify(properties)
    ? "same"
    : "occupied";
};

type ParagraphChain = {
  node: PMNode;
  /** The surviving paragraph: the chain's last, whose mark stays. */
  formattingOwnerPosition: number;
  chunks: (
    | { type: "paragraph"; source: PMNode; position: number }
    | { type: "bookmarks"; content: Fragment }
  )[];
  position: number;
  before: PMNode;
  joinBoundaries: { closing: number; bookmarks: Fragment }[];
};
type ResolveStructureOptions = {
  node: PMNode;
  position: number;
  context: StructuralContext;
};
const resolveStructure = ({
  node,
  position,
  context,
}: ResolveStructureOptions): PMNode | null => {
  if (node.isLeaf) return node;
  const textBoxAnchor = anchoredTextBoxId(node);
  if (
    textBoxAnchor !== undefined &&
    context.droppedTextBoxAnchors.has(textBoxAnchor)
  ) {
    context.structural = true;
    return null;
  }
  const structuralStart = {
    steps: context.steps.length,
    deleted: context.deleted.length,
    replacements: context.replacements.length,
    tableRanges: context.tableRanges.length,
    bookmarkJoins: context.bookmarkJoins.length,
  };
  const entries: { node: PMNode | null; original: PMNode; position: number }[] =
    [];
  node.forEach((child, offset) =>
    entries.push({
      node: resolveStructure({
        node: child,
        position: position + 1 + offset,
        context,
      }),
      original: child,
      position: position + 1 + offset,
    }),
  );
  const reversed: PMNode[] = [];
  let chain: ParagraphChain | null = null;
  let pendingBookmarks: PMNode[] = [];
  let followingContainerChild = false;
  const precedingParagraphs: boolean[] = [];
  let previousSurvivorIsParagraph = false;
  for (const { node: child } of entries) {
    precedingParagraphs.push(previousSurvivorIsParagraph);
    if (child !== null && !isBlockBoundaryNode(child)) {
      previousSurvivorIsParagraph = child.type.name === "paragraph";
    }
  }
  const flush = (): void => {
    if (!chain) return;
    let final = chain.node;
    if (chain.chunks.length > 1) {
      const children: PMNode[] = [];
      for (let index = chain.chunks.length - 1; index >= 0; index--) {
        const chunk = chain.chunks[index];
        if (!chunk) continue;
        // The surviving owner keeps its own marks, as the single-change join
        // does. Rebuilding them would change provenance without a style change.
        let content =
          chunk.type === "paragraph" ? chunk.source.content : chunk.content;
        if (
          chunk.type === "paragraph" &&
          chunk.position !== chain.formattingOwnerPosition
        ) {
          content = rebaseParagraphRunContent({
            paragraph: chunk.source,
            target: chain.node,
            position: chunk.position,
            styleResolver: context.styleResolver,
            onRebased: ({ before, after, position: runPosition }) =>
              recordNodeResolution({
                before,
                after,
                position: runPosition,
                steps: context.steps,
              }),
          });
        }
        content.forEach((child) => children.push(child));
      }
      final = rebuild(chain.node, { content: Fragment.fromArray(children) });
    }
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
    for (const { closing, bookmarks } of chain.joinBoundaries) {
      if (bookmarks.size === 0) {
        context.deleted.push({ from: closing, to: closing + 2 });
        continue;
      }
      // Higher-address deletions run first. Replace just the two paragraph
      // tokens and retained bookmarks, not any already-removed blocks between.
      const to = closing + bookmarks.size + 2;
      context.replacements.push({
        from: closing,
        to,
        slice: new Slice(bookmarks, 0, 0),
      });
      context.bookmarkJoins.push({
        from: closing,
        to,
        positionMap: new StepMap([0, 1, 0, bookmarks.size + 1, 1, 0]),
      });
    }
    reversed.push(
      chain.chunks.length === 1 && final.eq(chain.before)
        ? chain.before
        : final,
    );
    chain = null;
  };
  const flushBookmarks = (): void => {
    flush();
    for (const bookmark of pendingBookmarks) reversed.push(bookmark);
    pendingBookmarks = [];
  };
  for (let index = entries.length - 1; index >= 0; index--) {
    const entry = entries[index];
    if (!entry) continue;
    let paragraph = entry.node;
    if (paragraph === null) {
      // Break the loop's inference cycle through filler, paragraph, and chain.
      const hasSibling: boolean =
        index + reversed.length + pendingBookmarks.length + (chain ? 1 : 0) > 0;
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
      if (isBlockBoundaryNode(paragraph)) {
        pendingBookmarks.push(paragraph);
        continue;
      }
      flushBookmarks();
      reversed.push(paragraph);
      followingContainerChild ||= paragraph.type.spec["tableRole"] === "table";
      continue;
    }
    const marker = expectParagraphAttrs(paragraph).pPrMark;
    const markWasAdded = marker?.kind === "ins" || marker?.kind === "moveTo";
    const joins = marker && markWasAdded !== (context.mode === "accept");
    if (joins && chain) {
      const properties = expectParagraphAttrs(paragraph)._sectionProperties;
      const preceding = entries[index - 1]?.node;
      const outcome = paragraphMarkOutcome({
        follows: "paragraph",
        holdsContent: paragraphContentOf(paragraph),
        markWasAdded,
        containerStillEndsWithParagraph: true,
        followingTableCanReceiveRunIn: false,
        removedSectionProperties: markWasAdded && properties !== undefined,
        precedingSectionOwner: precedingSectionOwner(preceding, properties),
        followingHasSectionProperties: ownsSection(chain.node),
      });
      if (
        outcome.sectionDestination === "preceding" &&
        preceding?.type.name === "paragraph"
      ) {
        entries[index - 1]!.node = rebuild(preceding, {
          attrs: { ...preceding.attrs, _sectionProperties: properties },
        });
      } else if (outcome.sectionDestination === "following") {
        chain.node = rebuild(chain.node, {
          attrs: { ...chain.node.attrs, _sectionProperties: properties },
        });
      }
      const boundaries = pendingBookmarks.toReversed();
      let bookmarks = Fragment.empty;
      if (boundaries.length > 0) {
        const inline = inlineBookmarksOf(boundaries);
        if (!inline) {
          context.failed = true;
          return null;
        }
        bookmarks = Fragment.fromArray(inline);
        chain.chunks.push({ type: "bookmarks", content: bookmarks });
        pendingBookmarks = [];
      }
      // The mark that stays ends the joined paragraph, so the paragraph left
      // is the chain's last one: its identity, properties, mark and section
      // endpoint survive whether this one still has words or not. Its runs
      // are re-read in the survivor's style when the chain is assembled.
      const displacedToken = getProseParagraphPropertySourceToken(paragraph);
      const selectedToken = getProseParagraphPropertySourceToken(chain.node);
      context.transfers.push({
        displacedToken:
          typeof displacedToken === "string" ? displacedToken : null,
        selectedToken: typeof selectedToken === "string" ? selectedToken : null,
      });
      chain.chunks.push({
        type: "paragraph",
        source: paragraph,
        position: entry.position,
      });
      chain.position = entry.position;
      chain.before = paragraph;
      chain.joinBoundaries.push({
        closing: entry.position + paragraph.nodeSize - 1,
        bookmarks,
      });
      removedEndpoint(paragraph, context);
      continue;
    }
    flushBookmarks();
    let follows: "paragraph" | "table" | "nothing" = "nothing";
    if (chain) follows = "paragraph";
    else if (followingContainerChild) follows = "table";
    // With a table after it, the container must still end with a paragraph;
    // as the container's final paragraph, the block before it becomes last,
    // so it must be a paragraph (a cell or body never ends with a table).
    let containerStillEndsWithParagraph = true;
    if (follows === "table") {
      containerStillEndsWithParagraph =
        node.lastChild?.type.name === "paragraph";
    } else if (follows === "nothing") {
      containerStillEndsWithParagraph = precedingParagraphs[index] === true;
    }
    const outcome = paragraphMarkOutcome({
      follows,
      holdsContent: paragraphContentOf(paragraph),
      markWasAdded,
      containerStillEndsWithParagraph,
      followingTableCanReceiveRunIn: false,
      removedSectionProperties: markWasAdded && ownsSection(paragraph),
      precedingSectionOwner: precedingSectionOwner(
        entries[index - 1]?.node,
        expectParagraphAttrs(paragraph)._sectionProperties,
      ),
      followingHasSectionProperties: chain !== null && ownsSection(chain.node),
    });
    // A paragraph that goes leaves its range boundaries where it was; its
    // comment anchors go to the paragraph before it, and with none there to
    // take them it stays.
    const removes =
      marker &&
      !paragraphContentOf(paragraph) &&
      outcome.kind === "remove" &&
      joins &&
      index + reversed.length + 1 > 1;
    const remains = removes ? paragraphRemainsOf(paragraph) : undefined;
    // Before a table, bookmarks and moves stand in the paragraph's place; a
    // container ends with a paragraph, so ending it every boundary goes back.
    const between = follows === "table" ? (remains?.blocks ?? []) : [];
    const toPrevious =
      follows === "table" ? (remains?.comments ?? []) : (remains?.inline ?? []);
    const commentHost = entries[index - 1]?.node;
    if (
      remains !== undefined &&
      (toPrevious.length === 0 || commentHost?.type.name === "paragraph")
    ) {
      const properties = expectParagraphAttrs(paragraph)._sectionProperties;
      const hostEntry = entries[index - 1];
      if (toPrevious.length > 0 && commentHost && hostEntry) {
        const comments = Fragment.fromArray(toPrevious);
        hostEntry.node = rebuild(commentHost, {
          content: commentHost.content.append(comments),
        });
        const hostEnd = hostEntry.position + hostEntry.original.nodeSize - 1;
        context.replacements.push({
          from: hostEnd,
          to: hostEnd,
          slice: new Slice(comments, 0, 0),
        });
      }
      const preceding = entries[index - 1]?.node;
      if (
        outcome.sectionDestination === "preceding" &&
        preceding?.type.name === "paragraph"
      ) {
        entries[index - 1]!.node = rebuild(preceding, {
          attrs: { ...preceding.attrs, _sectionProperties: properties },
        });
      } else if (outcome.sectionDestination === "following" && chain) {
        chain.node = rebuild(chain.node, {
          attrs: { ...chain.node.attrs, _sectionProperties: properties },
        });
      }
      context.replacements.push({
        from: entry.position,
        to: entry.position + entry.original.nodeSize,
        slice: new Slice(Fragment.fromArray(between), 0, 0),
      });
      // The children are collected last to first.
      for (const block of between.toReversed()) reversed.push(block);
      removedEndpoint(paragraph, context);
      continue;
    }
    const resolved: PMNode = marker
      ? rebuild(paragraph, { attrs: { ...paragraph.attrs, pPrMark: null } })
      : paragraph;
    chain = {
      node: resolved,
      formattingOwnerPosition: entry.position,
      chunks: [
        { type: "paragraph", source: resolved, position: entry.position },
      ],
      position: entry.position,
      before: paragraph,
      joinBoundaries: [],
    };
    followingContainerChild = true;
  }
  flushBookmarks();
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
          context.deleted.push({
            from: entry.position,
            to: entry.position + carrier.nodeSize,
          });
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
  const unchanged =
    reversed.length === node.childCount &&
    reversed.every((child, index) => child === node.child(index));
  const rebuilt = unchanged ? node : rebuild(node, { content: reversed });
  if (node.type.spec["tableRole"] !== "table") return rebuilt;
  const tableResult = resolveAllTableChanges({
    table: rebuilt,
    mode: context.mode,
  });
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
      bookmarkJoins: context.bookmarkJoins.slice(structuralStart.bookmarkJoins),
      origin: position,
    });
    const positionMap = composeRevisionResolutionMaps(
      inputMap,
      tableResult.positionMap,
    );
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
  context.bookmarkJoins.length = structuralStart.bookmarkJoins;
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
  numbering: NumberingMap | null;
};
export const resolveWholeStory = ({
  doc,
  mode,
  styleResolver,
  numbering,
}: ResolveWholeStoryOptions) => {
  const { runsIntoTable, ...direct } = resolveStoryRevisions({
    doc,
    mode,
    styleResolver,
    numbering,
  });
  if (!runsIntoTable) return direct;
  // A paragraph whose mark goes right before a table runs on into the table's
  // first cell before the rest resolves; both read as one resolution.
  const runIn = new Transform(doc);
  const { targets, maps, transfers } = runParagraphsIntoTables(
    runIn,
    mode,
    styleResolver,
  );
  if (!runIn.docChanged) return direct;
  const { runsIntoTable: _rerun, ...result } = resolveStoryRevisions({
    doc: runIn.doc,
    mode,
    styleResolver,
    numbering,
  });
  for (const { position, step } of targets) {
    const from = runIn.mapping.slice(step).map(position);
    const target = runIn.doc.nodeAt(from);
    if (target) result.changedRanges.push({ from, to: from + target.nodeSize });
  }
  let runInMap = StepMap.empty;
  for (const map of maps)
    runInMap = composeRevisionResolutionMaps(runInMap, map);
  const positionMap = composeRevisionResolutionMaps(
    runInMap,
    result.positionMap,
  );
  return {
    ...result,
    steps: [...runIn.steps, ...result.steps],
    positionMap,
    mapping: new Mapping([positionMap]),
    transfers: [...transfers, ...result.transfers],
  };
};

const resolveStoryRevisions = ({
  doc,
  mode,
  styleResolver,
  numbering,
}: ResolveWholeStoryOptions) => {
  const steps: Step[] = [];
  const propertyRanges: { from: number; to: number }[] = [];
  const propertyResult = resolveProperties({
    doc,
    mode,
    styleResolver,
    numbering,
    steps,
    changedRanges: propertyRanges,
  });
  const properties = propertyResult.resolved;
  const inline = resolveInlineRevisions({
    pointComments: propertyResult.pointComments,
    doc: properties,
    mode,
    styleResolver,
    keepType:
      doc.type.schema.marks[mode === "accept" ? "insertion" : "deletion"],
    removeType:
      doc.type.schema.marks[mode === "accept" ? "deletion" : "insertion"],
  });
  const inlineResolved = inline?.resolved ?? properties;
  const comments = inline ? resolveCommentReferences(inlineResolved) : null;
  const context: StructuralContext = {
    mode,
    styleResolver,
    steps: [],
    changedRanges: [],
    deleted: [],
    replacements: [],
    transfers: [],
    tableRanges: [],
    bookmarkJoins: [],
    structural: propertyResult.structural,
    failed: false,
    removedEndpointCount: propertyResult.removedEndpointCount,
    removedReferences: [...propertyResult.removedReferences],
    droppedTextBoxAnchors: inline
      ? droppedTextBoxAnchorIds(properties, inlineResolved)
      : new Set<string>(),
  };
  const resolved =
    resolveStructure({
      node: comments?.doc ?? inlineResolved,
      position: -1,
      context,
    }) ?? doc;
  if (inline) steps.push(...inline.steps);
  if (comments) steps.push(...comments.steps);
  steps.push(...context.steps);
  const replacements = [
    ...context.replacements,
    ...context.deleted.map(({ from, to }) => ({
      from,
      to,
      slice: Slice.empty,
    })),
  ].sort((a, b) => a.from - b.from);
  const structuralMap = structuralReplacementMap({
    replacements,
    tables: context.tableRanges,
    bookmarkJoins: context.bookmarkJoins,
    origin: 0,
  });
  for (const { from, to, slice } of replacements.toReversed())
    steps.push(new ReplaceStep(from, to, slice));
  let preStructureMap = inline?.positionMap ?? StepMap.empty;
  for (const map of comments?.positionMaps ?? [])
    preStructureMap = composeRevisionResolutionMaps(preStructureMap, map);
  const positionMap = composeRevisionResolutionMaps(
    preStructureMap,
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
    preStructureMap,
    tableRanges: context.tableRanges,
    transfers: context.transfers,
    structuralRanges: context.changedRanges,
    structuralMap,
    structural: context.structural,
    failed: context.failed,
    removedEndpointCount: context.removedEndpointCount,
    removedReferences: context.removedReferences,
    runsIntoTable: propertyResult.runsIntoTable,
    /** The boundaries removed content left standing, for the range repair. */
    keptBoundaries: inline?.keptBoundaries ?? new Set<PMNode>(),
  };
};
