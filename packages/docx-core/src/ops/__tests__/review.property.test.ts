/**
 * The laws tracked operations and `resolveRevision` keep, over review-valid
 * synthetic documents (marks of every kind but none on a paragraph that ends
 * its container, deletions nested in insertions). Each tracked operation
 * carries a stamp dated after every change in the document, so it records
 * changes of its own rather than merging into one that was there.
 *
 * Documents are compared under π: equal once the fields a relayout
 * recomputes are dropped, adjacent records alike in everything but their
 * content and ids are merged (a container holding nothing stays: an empty
 * content control is markup of its own), and revision and content-control ids
 * are numbered in document order. π′ also ignores which paragraph's identity (id,
 * attributes, mark run properties) survives a join.
 *
 * - **L1 Accept.** Accepting a tracked operation's revisions gives the direct
 *   operation's result. A tracked insertion that splits another author's
 *   insertion, and a tracked deletion over deleted content, are compared once
 *   those changes are accepted on both sides too.
 * - **L2 Reject.** Rejecting them gives the document back, unless the
 *   operation kept a property change that was there before it. A split whose
 *   new paragraph is the second half gives it back under π′: removing the mark
 *   leaves the paragraph after it, which is the new one.
 * - **L3 Batches.** For a run of tracked operations, rejecting all their
 *   revisions at once equals rejecting each operation's in turn, from the
 *   last, and gives the document back. Accepting every revision in the story
 *   (A*) gives the same once the run's have been accepted operation by
 *   operation, in whatever order marks are removed, and rejecting every one
 *   (J*) gives what it gives on the document the run started from.
 * - **L4 Inverse.** Tracked operations and resolutions are undone exactly by
 *   their inverses, after a JSON round trip too.
 * - **L5 Determinism.** Equal inputs give equal results.
 * - **L6 Locality.** Paragraphs an operation does not touch are the same
 *   objects; it touches only paragraphs it names or resolves.
 * - **L7 Review shape.** No tracked operation leaves a mark on a paragraph
 *   that ends its container, and the seed contract holds.
 * - **Idempotence.** Resolving the same ids again changes nothing; resolved
 *   ids are gone.
 */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import {
  assertProperty,
  propertyConfig,
  propertyTestTimeout,
} from "../../../../../test/property-testing";
import type { BlockContent, Document, Paragraph, ParagraphFormatting } from "../../model/document";
import { applyDocumentOp, applyDocumentOps, type AppliedDocumentOp, stampOf } from "../apply";
import { storyParagraphs } from "../blocks";
import { contractViolation } from "../contract";
import { IDENTITY_SPACES, identityKeysIn, idKey, paragraphIdsIn } from "../ids";
import { gapAfterInserted } from "../inline";
import {
  asParagraphContent,
  childNodes,
  compareGaps,
  defaultInsertionGap,
  type Gap,
  type InlineNode,
  leafSpans,
  rebuildNode,
} from "../leaves";
import { paragraphLength, paragraphLogicalText } from "../offsets";
import { planTrackedDeletion, revisionIdDemand } from "../plan";
import { mergeAtSeam } from "../resolve";
import { DOCUMENT_OP_REFUSAL_REASONS } from "../refusal";
import { isTrackedWrapper, sameParagraphProperties, stampedRevisionIds } from "../review";
import {
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  type DocumentOpType,
  OP_STORIES,
  PARAGRAPH_MARK_FORMATTING_KEYS,
  REVISION_DECISIONS,
  type RevisionDecision,
  type RevisionStamp,
  SPLIT_HALVES,
} from "../types";
import {
  GENERATED_TRACKED_OP_KINDS,
  independentCopy,
  type OpSeed,
  opSeedArbitrary,
  reviewDocumentArbitrary,
  trackedOpFor,
} from "./documentArbitraries";

setDefaultTimeout(propertyTestTimeout(240_000));

/**
 * Each run applies an operation and checks it several ways, so the base count
 * is kept small: the nightly factor of ten takes it to 10^4, and the five-fold
 * run for changed areas stays within that job's time budget.
 */
const NUM_RUNS = 1000;

type Tally = Map<string, number>;

const count = (tally: Tally, key: string): void => {
  tally.set(key, (tally.get(key) ?? 0) + 1);
};

/**
 * A law that holds because nothing applied proves nothing. A replay of one
 * counterexample (`PROPERTY_TEST_PATH`) runs one example, which proves that
 * example and nothing about coverage.
 */
const expectEveryKindChecked = (tally: Tally, runs: number): void => {
  if (process.env["PROPERTY_TEST_PATH"] !== undefined) return;
  for (const kind of GENERATED_TRACKED_OP_KINDS) {
    expect({ kind, checked: (tally.get(kind) ?? 0) > runs / 2000 }).toEqual({
      kind,
      checked: true,
    });
  }
};

// ---------------------------------------------------------------------------
// π and π′
// ---------------------------------------------------------------------------

const canonicalList = (nodes: readonly InlineNode[]): InlineNode[] => {
  const out: InlineNode[] = [];
  for (const node of nodes) {
    const children = childNodes(node);
    const own = children === undefined ? node : rebuildNode(node, canonicalList(children));
    const last = out.at(-1);
    if (last === undefined) {
      out.push(own);
      continue;
    }
    // Merged as resolution merges records it leaves meeting; nothing here was emptied by it.
    out.splice(-1, 1, ...mergeAtSeam(last, own));
  }
  return out;
};

const MARK_KEYS: ReadonlySet<string> = new Set(PARAGRAPH_MARK_FORMATTING_KEYS);

const withoutMarkFormatting = (formatting: ParagraphFormatting): ParagraphFormatting =>
  // SAFETY: a subset of a property set's own entries.
  Object.fromEntries(
    Object.entries(formatting).filter(([key]) => !MARK_KEYS.has(key)),
  ) as ParagraphFormatting;

type Projection = "π" | "π′";

const canonicalParagraph = (paragraph: Paragraph, projection: Projection): Paragraph => {
  const next: Paragraph = {
    ...paragraph,
    content: asParagraphContent(canonicalList(paragraph.content)),
  };
  delete next.listRendering;
  delete next.renderedPageBreakBefore;
  if (projection === "π′") {
    delete next.paraId;
    delete next.textId;
    delete next.preservedAttributes;
    const formatting =
      next.formatting === undefined ? undefined : withoutMarkFormatting(next.formatting);
    if (formatting === undefined || Object.keys(formatting).length === 0) {
      delete next.formatting;
    } else {
      next.formatting = formatting;
    }
  }
  return next;
};

const canonicalBlocks = (blocks: readonly BlockContent[], projection: Projection): BlockContent[] =>
  blocks.map((block): BlockContent => {
    switch (block.type) {
      case "paragraph":
        return canonicalParagraph(block, projection);
      case "table":
        return {
          ...block,
          rows: block.rows.map((row) => ({
            ...row,
            cells: row.cells.map((cell) => ({
              ...cell,
              content: canonicalBlocks(cell.content, projection),
            })),
          })),
        };
      case "blockSdt":
      case "blockCustomXml":
        return { ...block, content: canonicalBlocks(block.content, projection) };
      default:
        return block;
    }
  });

/** Revision and content-control ids numbered in document order. */
const renumbered = (value: unknown, ids: Map<string, number>): unknown => {
  if (Array.isArray(value)) return value.map((item) => renumbered(item, ids));
  if (typeof value !== "object" || value === null) return value;
  const out: Record<string, unknown> = {};
  for (const [key, field] of Object.entries(value)) out[key] = renumbered(field, ids);
  const number = (key: string): number => {
    const known = ids.get(key);
    if (known !== undefined) return known;
    ids.set(key, ids.size + 1);
    return ids.size;
  };
  const info = out["info"];
  if (
    typeof info === "object" &&
    info !== null &&
    typeof Reflect.get(info, "author") === "string"
  ) {
    out["info"] = { ...info, id: number(`r:${String(Reflect.get(info, "id"))}`) };
  }
  if (typeof out["sdtType"] === "string" && typeof out["id"] === "number") {
    out["id"] = number(`c:${out["id"]}`);
  }
  return out;
};

const projected = (document: Document, projection: Projection = "π"): unknown =>
  renumbered(canonicalBlocks(document.package.document.content, projection), new Map());

const expectEquivalent = (
  actual: Document,
  expected: Document,
  projection: Projection = "π",
): void => {
  expect(projected(actual, projection)).toStrictEqual(projected(expected, projection));
};

// ---------------------------------------------------------------------------
// Operations and their revisions
// ---------------------------------------------------------------------------

const resolveOp = (ids: readonly number[], decision: RevisionDecision): DocumentOp => ({
  type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
  story: OP_STORIES.MAIN,
  revisionIds: ids,
  decision,
});

/** Resolve ids a law expects to resolve; a refusal is a failure. */
const resolved = (
  document: Document,
  ids: readonly number[],
  decision: RevisionDecision,
): Document => {
  const present = revisionIdsIn(document);
  const wanted = ids.filter((id) => present.has(id));
  if (wanted.length === 0) return document;
  const applied = applyDocumentOp(document, resolveOp(wanted, decision));
  if (applied.isErr()) throw applied.error;
  return applied.value.document;
};

/** Groups of ids resolved one group after another. */
const resolvedInTurn = (
  document: Document,
  groups: readonly (readonly number[])[],
  decision: RevisionDecision,
): Document => {
  let current = document;
  for (const ids of groups) current = resolved(current, ids, decision);
  return current;
};

/**
 * Every revision in the story resolved one way (A* or J*); `undefined` when
 * resolution refuses a change it does not carry out yet (a join at a section
 * break, a paragraph that would go).
 */
const resolvedAll = (document: Document, decision: RevisionDecision): Document | undefined => {
  const ids = [...storyRevisionIds(document)];
  if (ids.length === 0) return document;
  const applied = applyDocumentOp(document, resolveOp(ids, decision));
  if (applied.isErr()) {
    expect(applied.error.reason).toBe(DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE);
    return undefined;
  }
  return applied.value.document;
};

const revisionIdsIn = (value: unknown): Set<number> => {
  const prefix = `${IDENTITY_SPACES.REVISION}:`;
  return new Set(
    identityKeysIn(value).flatMap((key) =>
      key.startsWith(prefix) ? [Number(key.slice(prefix.length))] : [],
    ),
  );
};

const storyRevisionIds = (document: Document): Set<number> =>
  revisionIdsIn(document.package.document.content);

/** The revision ids of the records in a document's story carrying any of the stamps. */
const stampedIds = (document: Document, stamps: readonly RevisionStamp[]): number[] => {
  const paragraphs = storyParagraphs(document.package.document).map(({ paragraph }) => paragraph);
  return stamps.flatMap((stamp) => stampedRevisionIds(paragraphs, stamp, new Set()));
};

/**
 * The operation without its stamp: what it does directly. A direct join
 * merges no records, as a tracked one leaves that to resolution, which
 * merges as far as they are alike.
 */
const directOf = (op: DocumentOp): DocumentOp => {
  const direct = op.type === DOCUMENT_OP_TYPES.JOIN_BLOCKS ? { ...op, depth: 0 } : { ...op };
  Reflect.deleteProperty(direct, "revision");
  return direct;
};

const paragraphById = (document: Document, id: string): Paragraph | undefined =>
  storyParagraphs(document.package.document).find(
    ({ paragraph }) => idKey(paragraph.paraId ?? "") === idKey(id),
  )?.paragraph;

const follower = (document: Document, id: string): Paragraph | undefined => {
  const paragraphs = storyParagraphs(document.package.document);
  const at = paragraphs.find(({ paragraph }) => idKey(paragraph.paraId ?? "") === idKey(id));
  if (at === undefined) return undefined;
  return paragraphs.find(
    (candidate) =>
      candidate.index === at.index + 1 &&
      JSON.stringify(candidate.list) === JSON.stringify(at.list),
  )?.paragraph;
};

/** The leaves of a paragraph between two gaps. */
const leavesBetween = (paragraph: Paragraph, from: Gap, to: Gap) =>
  leafSpans(paragraph.content).filter((span) => {
    const start = compareGaps(span.before, from) >= 0 ? span.before : from;
    const end = compareGaps(span.after, to) <= 0 ? span.after : to;
    return compareGaps(start, end) < 0;
  });

/** The gaps an insertion's content occupies in the direct result. */
const insertedGaps = (op: DocumentOp, original: Paragraph): { from: Gap; to: Gap } | undefined => {
  switch (op.type) {
    case DOCUMENT_OP_TYPES.INSERT_TEXT: {
      return {
        from: { offset: op.at.offset, zeroWidthBefore: Number.MAX_SAFE_INTEGER },
        to: { offset: op.at.offset + op.text.length, zeroWidthBefore: 0 },
      };
    }
    case DOCUMENT_OP_TYPES.INSERT_CONTENT: {
      const from =
        op.at.zeroWidthBefore === undefined
          ? defaultInsertionGap(original.content, op.at.offset)
          : { offset: op.at.offset, zeroWidthBefore: op.at.zeroWidthBefore };
      return { from, to: gapAfterInserted(from, op.slice.content) };
    }
    default:
      return undefined;
  }
};

/**
 * Changes both sides of L1 accept besides the operation's own: another
 * author's insertions or moves the insertion lands in (and the pieces the
 * tracked one split them into), or deletions and moves away the deletion's
 * range overlaps. `undefined` when the law is not stated for the operation.
 */
const alsoAccepted = (
  document: Document,
  op: DocumentOp,
  direct: Document,
  tracked: AppliedDocumentOp,
): { direct: number[]; tracked: number[] } | undefined => {
  const original = revisionIdsIn(document.package.document.content);
  switch (op.type) {
    case DOCUMENT_OP_TYPES.INSERT_TEXT:
    case DOCUMENT_OP_TYPES.INSERT_CONTENT: {
      // Tracked changes the content brings get new ids on each side independently,
      // and their open ends can continue a change around the point.
      if (
        op.type === DOCUMENT_OP_TYPES.INSERT_CONTENT &&
        revisionIdsIn(op.slice.content).size > 0
      ) {
        return undefined;
      }
      const before = paragraphById(document, op.at.blockId);
      const after = paragraphById(direct, op.at.blockId);
      const gaps = before === undefined ? undefined : insertedGaps(op, before);
      if (after === undefined || gaps === undefined) return { direct: [], tracked: [] };
      const enclosing = new Set<number>();
      for (const span of leavesBetween(after, gaps.from, gaps.to)) {
        for (const ancestor of span.ancestors) {
          if (
            isTrackedWrapper(ancestor) &&
            (ancestor.type === "insertion" || ancestor.type === "moveTo") &&
            original.has(ancestor.info.id)
          ) {
            enclosing.add(ancestor.info.id);
          }
        }
      }
      if (enclosing.size === 0) return { direct: [], tracked: [] };
      const trackedParagraph = paragraphById(tracked.document, op.at.blockId);
      const pieces = new Set<number>();
      for (const { ancestors } of leafSpans(trackedParagraph?.content ?? [])) {
        for (const ancestor of ancestors) {
          if (
            isTrackedWrapper(ancestor) &&
            (ancestor.type === "insertion" || ancestor.type === "moveTo") &&
            !original.has(ancestor.info.id) &&
            !tracked.revisions.includes(ancestor.info.id)
          ) {
            pieces.add(ancestor.info.id);
          }
        }
      }
      return { direct: [...enclosing], tracked: [...enclosing, ...pieces] };
    }
    case DOCUMENT_OP_TYPES.DELETE_RANGE: {
      const paragraph = paragraphById(document, op.from.blockId);
      if (paragraph === undefined) return { direct: [], tracked: [] };
      const from = {
        offset: op.from.offset,
        zeroWidthBefore:
          op.from.zeroWidthBefore ??
          leafSpans(paragraph.content).filter(
            ({ before, after }) =>
              before.offset === op.from.offset && after.offset === op.from.offset,
          ).length,
      };
      const to = { offset: op.to.offset, zeroWidthBefore: op.to.zeroWidthBefore ?? 0 };
      const removed = new Set<number>();
      for (const span of leavesBetween(paragraph, from, to)) {
        // An empty deletion is a leaf of its own; the others hold the leaf.
        for (const record of [...span.ancestors, span.node]) {
          if (
            isTrackedWrapper(record) &&
            (record.type === "deletion" || record.type === "moveFrom")
          ) {
            removed.add(record.info.id);
          }
        }
      }
      return { direct: [...removed], tracked: [...removed] };
    }
    default:
      return { direct: [], tracked: [] };
  }
};

/**
 * Whether a tracked operation kept a property change its target already
 * carried: rejecting its own revisions then leaves its edit in place.
 */
const keptEarlierChange = (
  document: Document,
  op: DocumentOp,
  known: ReadonlySet<number>,
): boolean => {
  const earlier = (changes: readonly { info: { id: number } }[] | undefined): boolean =>
    (changes ?? []).some((change) => !known.has(change.info.id));
  switch (op.type) {
    case DOCUMENT_OP_TYPES.SET_RUN_PROPS: {
      const paragraph = paragraphById(document, op.from.blockId);
      if (paragraph === undefined) return false;
      const from = { offset: op.from.offset, zeroWidthBefore: 0 };
      const to = { offset: op.to.offset, zeroWidthBefore: 0 };
      return leavesBetween(paragraph, from, to).some(({ ancestors }) =>
        ancestors.some((ancestor) => ancestor.type === "run" && earlier(ancestor.propertyChanges)),
      );
    }
    case DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS:
      return earlier(paragraphById(document, op.blockId)?.propertyChanges);
    case DOCUMENT_OP_TYPES.JOIN_BLOCKS: {
      const next = paragraphById(document, op.nextBlockId);
      const first = paragraphById(document, op.blockId);
      return (
        first !== undefined &&
        paragraphLength(first) > 0 &&
        !sameParagraphProperties(first.formatting, next?.formatting) &&
        earlier(next?.propertyChanges)
      );
    }
    case DOCUMENT_OP_TYPES.SPLIT_BLOCK: {
      // A new second half takes the pending change, and with it the formatting it started from.
      const paragraph = paragraphById(document, op.at.blockId);
      if (paragraph === undefined || !newSecondHalf(paragraph, op)) return false;
      const fields = op.newParagraph ?? { formatting: paragraph.formatting };
      return (
        !sameParagraphProperties(fields.formatting, paragraph.formatting) &&
        earlier(paragraph.propertyChanges)
      );
    }
    default:
      return false;
  }
};

/** Whether a split's new paragraph is its second half: then rejecting it leaves the new id. */
const newSecondHalf = (paragraph: Paragraph, op: DocumentOp): boolean =>
  op.type === DOCUMENT_OP_TYPES.SPLIT_BLOCK &&
  (op.newHalf ??
    (op.at.offset === paragraphLength(paragraph) ? SPLIT_HALVES.SECOND : SPLIT_HALVES.FIRST)) ===
    SPLIT_HALVES.SECOND;

/**
 * The projection L2 compares a rejected tracked operation under: π′ for a
 * split whose new paragraph is the second half, since removing its mark
 * leaves that paragraph, under its new id.
 */
const rejectProjection = (document: Document, op: DocumentOp): Projection => {
  if (op.type !== DOCUMENT_OP_TYPES.SPLIT_BLOCK) return "π";
  const paragraph = paragraphById(document, op.at.blockId);
  return paragraph !== undefined && newSecondHalf(paragraph, op) ? "π′" : "π";
};

// ---------------------------------------------------------------------------
// Shape checks
// ---------------------------------------------------------------------------

/** Paragraphs carrying a mark that end their story body or table cell. */
const containerFinalMarks = (document: Document): string[] => {
  const out: string[] = [];
  const visit = (list: readonly BlockContent[], endsContainer: boolean): void => {
    for (const [index, block] of list.entries()) {
      const last = index === list.length - 1;
      switch (block.type) {
        case "paragraph":
          if (last && endsContainer && block.pPrMark !== undefined) out.push(block.paraId ?? "");
          break;
        case "table":
          for (const row of block.rows) for (const cell of row.cells) visit(cell.content, true);
          break;
        case "blockSdt":
        case "blockCustomXml":
          visit(block.content, last && endsContainer);
          break;
        default:
          break;
      }
    }
  };
  visit(document.package.document.content, true);
  return out;
};

const applyAll = (document: Document, ops: readonly DocumentOp[]): Document => {
  const applied = applyDocumentOps(document, ops);
  if (applied.isErr()) throw applied.error;
  return applied.value.document;
};

/** L4: the inverse restores exactly, after a JSON round trip too, and redoes. */
const expectRestores = (applied: AppliedDocumentOp, original: Document): void => {
  expect(contractViolation(applied.document)).toBeUndefined();
  const restored = applyDocumentOps(applied.document, applied.inverse);
  if (restored.isErr()) throw restored.error;
  expect(restored.value.document).toStrictEqual(original);
  // SAFETY: operations are plain data; this is the journal's round trip.
  const replayed = JSON.parse(JSON.stringify(applied.inverse)) as DocumentOp[];
  expect(applyAll(applied.document, replayed)).toStrictEqual(original);
  expect(applyAll(restored.value.document, restored.value.inverse)).toStrictEqual(applied.document);
};

const outcome = (result: ReturnType<typeof applyDocumentOp>) =>
  result.isOk() ? result.value : { refused: result.error.reason };

/** Revision ids to resolve drawn from a document: a random subset of those its story holds. */
const drawIds = (document: Document, picks: readonly boolean[]): number[] =>
  [...storyRevisionIds(document)].filter((_, index) => picks[index % Math.max(1, picks.length)]);

const decisionArbitrary = fc.constantFrom(REVISION_DECISIONS.ACCEPT, REVISION_DECISIONS.REJECT);

const kindOf = (op: DocumentOp): DocumentOpType => op.type;

/** Why a tracked operation is refused where the direct one applies. */
const TRACKING_REFUSALS: readonly string[] = [
  DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
  DOCUMENT_OP_REFUSAL_REASONS.REVISION_CONFLICT,
  DOCUMENT_OP_REFUSAL_REASONS.INSIDE_TRACKED_DELETION,
  // A tracked change is a record more, whose id a cut may need to name.
  DOCUMENT_OP_REFUSAL_REASONS.NEEDS_NEW_IDS,
];

// ---------------------------------------------------------------------------
// The laws
// ---------------------------------------------------------------------------

describe("tracked operations and their resolution", () => {
  test("L1: accepting a tracked operation's revisions gives the direct operation's result", () => {
    const tally: Tally = new Map();
    assertProperty(
      fc.property(reviewDocumentArbitrary, opSeedArbitrary, (document, seed) => {
        const op = trackedOpFor(document, seed);
        const tracked = applyDocumentOp(document, op);
        const direct = applyDocumentOp(document, directOf(op));
        if (tracked.isErr() && direct.isOk()) {
          // Tracking refuses only what no tracked change can record.
          expect(TRACKING_REFUSALS).toContain(tracked.error.reason);
        }
        if (tracked.isErr() || direct.isErr()) {
          count(tally, "refused");
          return;
        }
        const extra = alsoAccepted(document, op, direct.value.document, tracked.value);
        if (extra === undefined) {
          count(tally, "notStated");
          return;
        }
        count(tally, kindOf(op));
        const accepted = resolved(
          tracked.value.document,
          [...tracked.value.revisions, ...extra.tracked],
          REVISION_DECISIONS.ACCEPT,
        );
        const expected = resolved(direct.value.document, extra.direct, REVISION_DECISIONS.ACCEPT);
        expectEquivalent(accepted, expected);
      }),
      { numRuns: NUM_RUNS },
    );
    expectEveryKindChecked(tally, NUM_RUNS);
  });

  test("L2: rejecting a tracked operation's revisions gives the document back", () => {
    const tally: Tally = new Map();
    fc.assert(
      fc.property(reviewDocumentArbitrary, opSeedArbitrary, (document, seed) => {
        const op = trackedOpFor(document, seed);
        const tracked = applyDocumentOp(document, op);
        if (tracked.isErr()) {
          count(tally, `refused ${op.type} ${tracked.error.reason}`);
          return;
        }
        if (keptEarlierChange(document, op, new Set())) {
          count(tally, "keptEarlierChange");
          return;
        }
        count(tally, kindOf(op));
        const rejected = resolved(
          tracked.value.document,
          tracked.value.revisions,
          REVISION_DECISIONS.REJECT,
        );
        expectEquivalent(rejected, document, rejectProjection(document, op));
      }),
      propertyConfig({ numRuns: NUM_RUNS }),
    );
    expectEveryKindChecked(tally, NUM_RUNS);
  });

  test("L3: resolving a run's revisions at once equals resolving each operation's in turn", () => {
    const tally: Tally = new Map();
    fc.assert(
      fc.property(
        reviewDocumentArbitrary,
        fc.array(opSeedArbitrary, { minLength: 2, maxLength: 6 }),
        (document, seeds: OpSeed[]) => {
          let current = document;
          const stamps: RevisionStamp[] = [];
          let keptEarlier = false;
          for (const [index, seed] of seeds.entries()) {
            const op = trackedOpFor(current, seed, index);
            const applied = applyDocumentOp(current, op);
            if (applied.isErr()) continue;
            const stamp = stampOf(op);
            if (stamp === undefined) continue;
            keptEarlier ||= keptEarlierChange(current, op, new Set(stampedIds(current, stamps)));
            count(tally, kindOf(op));
            stamps.push(stamp);
            current = applied.value.document;
          }
          // An operation's revisions: every record carrying its stamp, pieces later
          // operations cut from them included.
          const perOp = stamps.map((stamp) => stampedIds(current, [stamp]));
          const all = perOp.flat();
          if (all.length === 0) return;
          // The run's own revisions: at once, and one operation's at a time.
          const rejectedAtOnce = resolved(current, all, REVISION_DECISIONS.REJECT);
          const rejectedInTurn = resolvedInTurn(
            current,
            perOp.toReversed(),
            REVISION_DECISIONS.REJECT,
          );
          expectEquivalent(rejectedAtOnce, rejectedInTurn, "π′");
          // A split at a paragraph's end is given back under the new paragraph's id.
          if (!keptEarlier) {
            expectEquivalent(rejectedAtOnce, document, "π′");
          }
          // Every revision (A*, J*): the run's in turn first, or all at once.
          const acceptedInTurn = resolvedInTurn(current, perOp, REVISION_DECISIONS.ACCEPT);
          const allAccepted = resolvedAll(current, REVISION_DECISIONS.ACCEPT);
          const turnAccepted = resolvedAll(acceptedInTurn, REVISION_DECISIONS.ACCEPT);
          if (allAccepted !== undefined && turnAccepted !== undefined) {
            count(tally, "A*");
            expectEquivalent(allAccepted, turnAccepted, "π′");
          }
          const allRejected = resolvedAll(current, REVISION_DECISIONS.REJECT);
          const originalRejected = resolvedAll(document, REVISION_DECISIONS.REJECT);
          if (allRejected !== undefined && originalRejected !== undefined) {
            count(tally, "J*");
            expectEquivalent(allRejected, originalRejected, "π′");
          }
        },
      ),
      propertyConfig({ numRuns: NUM_RUNS / 5 }),
    );
    expectEveryKindChecked(tally, NUM_RUNS / 5);
  });

  test("L4: tracked operations and resolutions are undone exactly by their inverses", () => {
    const tally: Tally = new Map();
    fc.assert(
      fc.property(
        reviewDocumentArbitrary,
        opSeedArbitrary,
        fc.array(fc.boolean(), { minLength: 1, maxLength: 6 }),
        decisionArbitrary,
        (document, seed, picks, decision) => {
          const original = structuredClone(document);
          const op = trackedOpFor(document, seed);
          const applied = applyDocumentOp(document, op);
          expect(document).toStrictEqual(original);
          if (applied.isOk()) {
            count(tally, kindOf(op));
            expectRestores(applied.value, original);
          }
          const ids = drawIds(document, picks);
          const resolution = applyDocumentOp(document, resolveOp(ids, decision));
          expect(document).toStrictEqual(original);
          if (resolution.isOk()) {
            count(tally, "resolved");
            expectRestores(resolution.value, original);
          } else {
            count(tally, `resolution ${resolution.error.reason}`);
          }
        },
      ),
      propertyConfig({ numRuns: NUM_RUNS }),
    );
    expectEveryKindChecked(tally, NUM_RUNS);
    expect(tally.get("resolved") ?? 0).toBeGreaterThan(NUM_RUNS / 4);
  });

  test("L4: a run of tracked operations is undone exactly, in reverse and as a batch", () => {
    const tally: Tally = new Map();
    fc.assert(
      fc.property(
        reviewDocumentArbitrary,
        fc.array(opSeedArbitrary, { minLength: 2, maxLength: 6 }),
        (document, seeds) => {
          const original = structuredClone(document);
          let current = document;
          const ops: DocumentOp[] = [];
          const inverses: (readonly DocumentOp[])[] = [];
          for (const [index, seed] of seeds.entries()) {
            const op = trackedOpFor(current, seed, index);
            const applied = applyDocumentOp(current, op);
            if (applied.isErr()) continue;
            count(tally, kindOf(op));
            ops.push(op);
            inverses.push(applied.value.inverse);
            current = applied.value.document;
          }
          expect(applyAll(current, inverses.toReversed().flat())).toStrictEqual(original);
          const batch = applyDocumentOps(document, ops);
          if (batch.isErr()) throw batch.error;
          expect(batch.value.document).toStrictEqual(current);
          expectRestores(batch.value, original);
        },
      ),
      propertyConfig({ numRuns: NUM_RUNS / 5 }),
    );
    expectEveryKindChecked(tally, NUM_RUNS / 5);
  });

  test("L5: equal inputs give equal results", () => {
    fc.assert(
      fc.property(
        reviewDocumentArbitrary,
        opSeedArbitrary,
        fc.array(fc.boolean(), { minLength: 1, maxLength: 6 }),
        decisionArbitrary,
        (document, seed, picks, decision) => {
          for (const op of [
            trackedOpFor(document, seed),
            resolveOp(drawIds(document, picks), decision),
          ]) {
            // SAFETY: the operation is plain data; this is the journal's round trip.
            const replayed = JSON.parse(JSON.stringify(op)) as DocumentOp;
            const first = outcome(applyDocumentOp(structuredClone(document), op));
            expect(outcome(applyDocumentOp(document, replayed))).toStrictEqual(first);
            expect(outcome(applyDocumentOp(independentCopy(document), replayed))).toStrictEqual(
              first,
            );
          }
        },
      ),
      propertyConfig({ numRuns: NUM_RUNS }),
    );
  });

  test("L6: an operation touches only the paragraphs it names or resolves", () => {
    fc.assert(
      fc.property(
        reviewDocumentArbitrary,
        opSeedArbitrary,
        fc.array(fc.boolean(), { minLength: 1, maxLength: 6 }),
        decisionArbitrary,
        (document, seed, picks, decision) => {
          const ids = drawIds(document, picks);
          const cases: [DocumentOp, Set<string>][] = [];
          const op = trackedOpFor(document, seed);
          cases.push([op, namedParagraphs(op)]);
          const holding = new Set<string>();
          for (const { paragraph } of storyParagraphs(document.package.document)) {
            const own = revisionIdsIn(paragraph);
            if (!ids.some((id) => own.has(id))) continue;
            holding.add(idKey(paragraph.paraId ?? ""));
            const next = follower(document, paragraph.paraId ?? "");
            if (next !== undefined) holding.add(idKey(next.paraId ?? ""));
          }
          cases.push([resolveOp(ids, decision), holding]);
          for (const [candidate, named] of cases) {
            const applied = applyDocumentOp(document, candidate);
            if (applied.isErr()) continue;
            const touched = new Set(
              [
                ...applied.value.touched.modified,
                ...applied.value.touched.inserted,
                ...applied.value.touched.removed,
              ].map(idKey),
            );
            for (const id of touched) expect(named.has(id)).toBe(true);
            const after = new Map(
              storyParagraphs(applied.value.document.package.document).map(({ paragraph }) => [
                idKey(paragraph.paraId ?? ""),
                paragraph,
              ]),
            );
            for (const { paragraph } of storyParagraphs(document.package.document)) {
              const key = idKey(paragraph.paraId ?? "");
              if (!touched.has(key)) expect(after.get(key)).toBe(paragraph);
            }
          }
        },
      ),
      propertyConfig({ numRuns: NUM_RUNS }),
    );
  });

  test("L7: no tracked operation marks a paragraph that ends its container", () => {
    const tally: Tally = new Map();
    fc.assert(
      fc.property(reviewDocumentArbitrary, opSeedArbitrary, (document, seed) => {
        expect(containerFinalMarks(document)).toEqual([]);
        const op = trackedOpFor(document, seed);
        const applied = applyDocumentOp(document, op);
        if (applied.isErr()) return;
        count(tally, kindOf(op));
        expect(containerFinalMarks(applied.value.document)).toEqual([]);
        expect(contractViolation(applied.value.document)).toBeUndefined();
      }),
      propertyConfig({ numRuns: NUM_RUNS }),
    );
    expectEveryKindChecked(tally, NUM_RUNS);
  });

  test("resolving the same revisions again changes nothing", () => {
    const tally: Tally = new Map();
    fc.assert(
      fc.property(
        reviewDocumentArbitrary,
        fc.array(fc.boolean(), { minLength: 1, maxLength: 6 }),
        decisionArbitrary,
        (document, picks, decision) => {
          const ids = drawIds(document, picks);
          const once = applyDocumentOp(document, resolveOp(ids, decision));
          if (once.isErr()) {
            count(tally, "refused");
            return;
          }
          count(tally, "resolved");
          const remaining = storyRevisionIds(once.value.document);
          for (const id of ids) expect(remaining.has(id)).toBe(false);
          const twice = applyDocumentOp(once.value.document, resolveOp(ids, decision));
          if (twice.isErr()) throw twice.error;
          expect(twice.value.document).toBe(once.value.document);
          expect(twice.value.inverse).toEqual([]);
        },
      ),
      propertyConfig({ numRuns: NUM_RUNS }),
    );
    expect(tally.get("resolved") ?? 0).toBeGreaterThan(NUM_RUNS / 4);
  });

  test("a tracked operation takes exactly the new ids revisionIdDemand counts", () => {
    fc.assert(
      fc.property(reviewDocumentArbitrary, opSeedArbitrary, (document, seed) => {
        const op = trackedOpFor(document, { ...seed, depth: 1 });
        const demand = revisionIdDemand(document, op);
        if (demand.isErr()) return;
        const base = 900_000;
        const withIds = (length: number): DocumentOp => ({
          ...op,
          newIds: {
            revision: Array.from({ length }, (_, index) => base + index),
            control: Array.from({ length: 64 }, (_, index) => base + index),
          },
        });
        expect(applyDocumentOp(document, withIds(demand.value)).isOk()).toBe(true);
        if (demand.value > 0) {
          const short = applyDocumentOp(document, withIds(demand.value - 1));
          expect(short.isErr() && short.error.reason).toBe(
            DOCUMENT_OP_REFUSAL_REASONS.NEEDS_NEW_IDS,
          );
        }
      }),
      propertyConfig({ numRuns: NUM_RUNS / 5 }),
    );
  });

  test("a planned tracked deletion removes only the author's own insertions", () => {
    const tally: Tally = new Map();
    assertProperty(
      fc.property(reviewDocumentArbitrary, opSeedArbitrary, (document, seed) => {
        const op = trackedOpFor(document, { ...seed, kind: 2 });
        if (op.type !== DOCUMENT_OP_TYPES.DELETE_RANGE || op.revision === undefined) return;
        const plan = planTrackedDeletion(document, {
          from: op.from,
          to: op.to,
          revision: op.revision,
          newIds: { revision: Array.from({ length: 64 }, (_, index) => 700_000 + index) },
        });
        if (plan.isErr()) {
          count(tally, "refused");
          return;
        }
        const applied = applyDocumentOps(document, plan.value);
        if (applied.isErr()) throw applied.error;
        count(tally, plan.value.length > 1 ? "several" : "one");
        const paragraph = paragraphById(document, op.from.blockId);
        if (paragraph === undefined) return;
        const text = paragraphLogicalText(paragraph);
        const from = {
          offset: op.from.offset,
          zeroWidthBefore: op.from.zeroWidthBefore ?? Number.MAX_SAFE_INTEGER,
        };
        const to = { offset: op.to.offset, zeroWidthBefore: op.to.zeroWidthBefore ?? 0 };
        const own = new Set<number>();
        for (const span of leavesBetween(paragraph, from, to)) {
          const removed = span.ancestors.some(
            (ancestor) => ancestor.type === "deletion" || ancestor.type === "moveFrom",
          );
          const mine = span.ancestors.some(
            (ancestor) =>
              ancestor.type === "insertion" &&
              isTrackedWrapper(ancestor) &&
              ancestor.info.author === op.revision?.author,
          );
          if (removed || !mine) continue;
          const start = Math.max(span.before.offset, op.from.offset);
          const end = Math.min(span.after.offset, op.to.offset);
          for (let unit = start; unit < end; unit += 1) own.add(unit);
        }
        const expected = Array.from({ length: text.length }, (_, unit) =>
          own.has(unit) ? "" : text.charAt(unit),
        ).join("");
        const after = paragraphById(applied.value.document, op.from.blockId);
        expect(after === undefined ? undefined : paragraphLogicalText(after)).toBe(expected);
        // Rejecting what the plan recorded leaves only the retraction.
        const retracted = applyAll(
          document,
          plan.value.filter(
            (planned) =>
              planned.type === DOCUMENT_OP_TYPES.DELETE_RANGE && planned.revision === undefined,
          ),
        );
        expectEquivalent(
          resolved(applied.value.document, applied.value.revisions, REVISION_DECISIONS.REJECT),
          retracted,
        );
      }),
      { numRuns: NUM_RUNS / 5 },
    );
    expect(tally.get("one") ?? 0).toBeGreaterThan(0);
  });
});

/** The paragraphs an operation names. */
const namedParagraphs = (op: DocumentOp): Set<string> => {
  switch (op.type) {
    case DOCUMENT_OP_TYPES.DELETE_BLOCKS:
      return new Set(op.blockIds.map(idKey));
    case DOCUMENT_OP_TYPES.INSERT_BLOCKS:
      return new Set([
        idKey(op.at.blockId),
        ...op.blocks.flatMap(({ paraId }) => (paraId === undefined ? [] : [idKey(paraId)])),
      ]);
    case DOCUMENT_OP_TYPES.INSERT_TEXT:
    case DOCUMENT_OP_TYPES.INSERT_CONTENT:
    case DOCUMENT_OP_TYPES.SPLIT_INLINE:
    case DOCUMENT_OP_TYPES.JOIN_INLINE:
      return new Set([idKey(op.at.blockId)]);
    case DOCUMENT_OP_TYPES.DELETE_RANGE:
    case DOCUMENT_OP_TYPES.SET_RUN_PROPS:
      return new Set([idKey(op.from.blockId), idKey(op.to.blockId)]);
    case DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS:
    case DOCUMENT_OP_TYPES.SET_PARAGRAPH_REVIEW:
    case DOCUMENT_OP_TYPES.REPLACE_INLINE:
      return new Set([idKey(op.blockId)]);
    case DOCUMENT_OP_TYPES.SPLIT_BLOCK:
      return new Set([idKey(op.at.blockId), idKey(op.newBlockId)]);
    case DOCUMENT_OP_TYPES.JOIN_BLOCKS:
      return new Set([idKey(op.blockId), idKey(op.nextBlockId)]);
    case DOCUMENT_OP_TYPES.REPLACE_BLOCKS:
      return new Set(
        [...op.expected, ...op.blocks].flatMap(({ paraId }) =>
          paraId === undefined ? [] : [idKey(paraId)],
        ),
      );
    case DOCUMENT_OP_TYPES.INSERT_TABLE:
      return new Set([idKey(op.at.blockId), ...paragraphIdsIn(op.table).map(idKey)]);
    case DOCUMENT_OP_TYPES.DELETE_TABLE:
      return new Set([idKey(op.blockId), ...paragraphIdsIn(op.expected ?? []).map(idKey)]);
    case DOCUMENT_OP_TYPES.SET_CONTAINER_BLOCKS:
      return new Set([idKey(op.blockId), ...paragraphIdsIn([op.expected, op.blocks]).map(idKey)]);
    case DOCUMENT_OP_TYPES.INSERT_ROW:
      return new Set([idKey(op.blockId), ...paragraphIdsIn(op.row).map(idKey)]);
    case DOCUMENT_OP_TYPES.DELETE_ROW:
      return new Set([idKey(op.blockId), ...paragraphIdsIn(op.expected ?? []).map(idKey)]);
    case DOCUMENT_OP_TYPES.SET_TABLE_ROWS:
      return new Set([idKey(op.blockId), ...paragraphIdsIn([op.expected, op.rows]).map(idKey)]);
    case DOCUMENT_OP_TYPES.RESOLVE_REVISION:
      return new Set();
    default: {
      const unreachable: never = op;
      return unreachable;
    }
  }
};
