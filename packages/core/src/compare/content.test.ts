import { describe, expect, test } from "bun:test";
import fc from "fast-check";

import { propertyConfig } from "../../../../test/property-testing";
import {
  compareContent,
  createContentComparisonWorkSession,
  FOLIO_CONTENT_COMPARISON_LIMITS,
  FolioContentComparisonLimitError,
  InvalidFolioContentComparisonError,
  type FolioContentComparison,
  type FolioContentComparisonEvent,
  type FolioContentPairRelation,
  type FolioContentTextSegment,
} from "./content";
import type {
  FolioContentBlock,
  FolioContentPropertyInput,
  FolioContentPropertyInputValue,
  FolioContentSnapshot,
} from "./content-types";

type TestBlockKind = "heading" | "paragraph";
type TestBlock = FolioContentBlock<TestBlockKind>;

type LegacyTestRun = {
  text: string;
  bold?: boolean;
  italic?: boolean;
  color?: string;
  directFormatting?: Readonly<Record<string, FolioContentPropertyInputValue>>;
};

type TestBlockOptions = {
  id: string;
  text: string;
  kind?: TestBlockKind;
  identitySemantics?: TestBlock["identity"]["type"];
  headingLevel?: number;
  displayLabel?: string;
  styleId?: string;
  listLevel?: number;
  directAlignment?: string;
  directSpacing?: Readonly<Record<string, FolioContentPropertyInputValue>>;
  effectiveParagraphFormatting?: Readonly<
    Record<string, FolioContentPropertyInputValue | undefined>
  >;
  runs?: readonly LegacyTestRun[];
  structuralBoundaries?: TestBlock["structuralBoundaries"];
  table?: TestBlock["table"];
  containerPath?: readonly {
    kind: string;
    id: string;
    identitySemantics?: TestBlock["identity"]["type"];
  }[];
};

const propertySet = (
  properties: Readonly<Record<string, FolioContentPropertyInputValue | undefined>>,
): FolioContentPropertyInput =>
  Object.entries(properties)
    .filter((entry): entry is [string, FolioContentPropertyInputValue] => entry[1] !== undefined)
    .map(([key, value]) => ({ key, value }))
    .toSorted((left, right) => (left.key < right.key ? -1 : left.key > right.key ? 1 : 0));

const objectProperty = (
  properties: Readonly<Record<string, FolioContentPropertyInputValue | undefined>>,
): FolioContentPropertyInputValue => ({ type: "object", entries: propertySet(properties) });

const contentBlock = ({
  id,
  text,
  kind = "paragraph",
  identitySemantics = "authoritative",
  headingLevel,
  displayLabel,
  styleId,
  listLevel,
  directAlignment,
  directSpacing,
  effectiveParagraphFormatting,
  runs = [],
  structuralBoundaries = [],
  table,
  containerPath = [],
}: TestBlockOptions): TestBlock => ({
  identity: { type: identitySemantics, id },
  kind,
  text,
  blockProperties: propertySet({ headingLevel, displayLabel }),
  paragraphFormatting: {
    authored: propertySet({
      styleId,
      listLevel,
      directAlignment,
      directSpacing: directSpacing === undefined ? undefined : objectProperty(directSpacing),
    }),
    effective: propertySet(effectiveParagraphFormatting ?? {}),
  },
  runs: runs.map(({ text: runText, directFormatting, ...effective }) => ({
    text: runText,
    effectiveFormatting: propertySet(effective),
    authoredFormatting: propertySet(directFormatting ?? {}),
  })),
  structuralBoundaries,
  ...(table !== undefined && { table }),
  containerPath: containerPath.map(
    ({ kind: containerKind, id: containerId, identitySemantics: containerIdentitySemantics }) => ({
      kind: containerKind,
      identity: {
        type: containerIdentitySemantics ?? "authoritative",
        id: containerId,
      },
    }),
  ),
});

type TableBlockOptions = {
  id: string;
  text: string;
  rowIndex: number;
  cellIndex: number;
  gridColumnIndex?: number;
  tableIndex?: number;
  columnSpan?: number;
  rowSpan?: number;
  paragraphIndex?: number;
  outerTableIdentity?: TestBlock["identity"];
  tableIdentity?: TestBlock["identity"];
  rowIdentity?: TestBlock["identity"];
  cellIdentity?: TestBlock["identity"];
  containerPath?: TestBlockOptions["containerPath"];
};

const tableBlock = ({
  id,
  text,
  rowIndex,
  cellIndex,
  gridColumnIndex = cellIndex,
  tableIndex = 0,
  columnSpan = 1,
  rowSpan = 1,
  paragraphIndex = 0,
  outerTableIdentity = { type: "positional", id: `outer-table-${String(tableIndex)}` },
  tableIdentity = { type: "positional", id: `table-${String(tableIndex)}` },
  rowIdentity = {
    type: "positional",
    id: `table-${String(tableIndex)}-row-${String(rowIndex)}`,
  },
  cellIdentity = {
    type: "positional",
    id: `table-${String(tableIndex)}-row-${String(rowIndex)}-cell-${String(cellIndex)}`,
  },
  containerPath,
}: TableBlockOptions): TestBlock =>
  contentBlock({
    id,
    text,
    table: {
      outerTableIdentity,
      tableIdentity,
      rowIdentity,
      cellIdentity,
      outerTableIndex: tableIndex,
      tableIndex,
      rowIndex,
      cellIndex,
      gridColumnIndex,
      columnSpan,
      rowSpan,
      paragraphIndex,
    },
    ...(containerPath !== undefined && { containerPath }),
  });

type SuccessfulComparisonOptions = {
  base: TestBlock[];
  revised: TestBlock[];
  granularity?: "word" | "character";
};

const successfulComparison = ({
  base,
  revised,
  granularity,
}: SuccessfulComparisonOptions): FolioContentComparison => {
  const result = compareContent({
    base: { blocks: base },
    revised: { blocks: revised },
    ...(granularity && { granularity }),
  });
  if (result.isErr()) {
    throw result.error;
  }
  expect(result.isErr()).toBe(false);
  return result.value;
};

const eventTypes = (comparison: FolioContentComparison): FolioContentComparisonEvent["type"][] =>
  comparison.events.map(({ type }) => type);

const changedBlockKeys = (relation: FolioContentPairRelation): string[] =>
  relation.blockChanges.flatMap((change) =>
    change.field === "blockProperties"
      ? change.changes.map(({ key }) => key)
      : [change.field],
  );

const structuralEventBlock = (
  event: Extract<FolioContentComparisonEvent, { type: "structural" }>,
): FolioContentBlock => {
  const block = event.change.blocks.at(event.memberIndex);
  if (!block) throw new Error("Structural comparison event has no owned member");
  return block;
};

const expectRecursivelyFrozen = (value: unknown, visited = new WeakSet<object>()): void => {
  if ((typeof value !== "object" && typeof value !== "function") || value === null) return;
  if (visited.has(value)) return;
  visited.add(value);
  expect(Object.isFrozen(value)).toBe(true);
  for (const key of Reflect.ownKeys(value)) {
    const descriptor = Object.getOwnPropertyDescriptor(value, key);
    if (descriptor && "value" in descriptor) {
      expectRecursivelyFrozen(descriptor.value, visited);
    }
  }
};

const assertComparisonResultIsDeeplyReadonly = (comparison: FolioContentComparison): void => {
  // @ts-expect-error comparison event arrays are readonly
  comparison.events.push({});
  const event = comparison.events.at(0);
  if (event?.type === "modified") {
    // @ts-expect-error captured identity leaves are readonly
    event.relation.base.block.identity.id = "mutated";
    // @ts-expect-error canonical segment leaves are readonly
    event.relation.segments[0].text = "mutated";
    // @ts-expect-error canonical property arrays are readonly
    event.relation.base.block.blockProperties.push({ key: "x", value: true });
    const propertyValue = event.relation.base.block.blockProperties.at(0)?.value;
    if (typeof propertyValue === "object" && propertyValue?.type === "object") {
      // @ts-expect-error nested canonical property arrays are readonly
      propertyValue.entries.push({ key: "x", value: true });
    }
  }
  if (event?.type === "structural") {
    // @ts-expect-error structural member tuples are readonly
    event.change.blocks[0] = event.change.blocks[0];
  }
  if (event?.type === "tableReplacement") {
    // @ts-expect-error replacement member tuples are readonly
    event.replacement.baseBlocks[0] = event.replacement.baseBlocks[0];
    // @ts-expect-error nested comparison events are readonly
    event.replacement.refinement.events.push({});
  }
};
void assertComparisonResultIsDeeplyReadonly;

const eventBaseBlocks = (event: FolioContentComparisonEvent): FolioContentBlock[] => {
  switch (event.type) {
    case "unchanged":
    case "modified":
    case "formatting":
      return [event.relation.base.block];
    case "deleted":
      return [event.block];
    case "movedFrom":
      return [event.move.relation.base.block];
    case "inserted":
      return [];
    case "movedTo":
      return [];
    case "split":
      return [event.relations[0].base.block];
    case "merge":
      return event.relations.map(({ base }) => base.block);
    case "tableReplacement":
      return [...event.replacement.baseBlocks];
    case "structural":
      return event.change.type === "table-delete" ||
        event.change.type === "table-row-delete" ||
        event.change.type === "table-column-delete"
        ? [structuralEventBlock(event)]
        : [];
    default: {
      const unreachable: never = event;
      throw new Error(`Unhandled event ${JSON.stringify(unreachable)}`);
    }
  }
};

const eventRevisedBlocks = (event: FolioContentComparisonEvent): FolioContentBlock[] => {
  switch (event.type) {
    case "unchanged":
    case "modified":
    case "formatting":
      return [event.relation.revised.block];
    case "movedTo":
      return [event.move.relation.revised.block];
    case "inserted":
      return [event.block];
    case "deleted":
    case "movedFrom":
      return [];
    case "split":
      return event.relations.map(({ revised }) => revised.block);
    case "merge":
      return [event.relations[0].revised.block];
    case "tableReplacement":
      return [...event.replacement.revisedBlocks];
    case "structural":
      return event.change.type === "table-insert" ||
        event.change.type === "table-row-insert" ||
        event.change.type === "table-column-insert"
        ? [structuralEventBlock(event)]
        : [];
    default: {
      const unreachable: never = event;
      throw new Error(`Unhandled event ${JSON.stringify(unreachable)}`);
    }
  }
};

const baseProjection = (comparison: FolioContentComparison): FolioContentBlock[] =>
  comparison.events.flatMap(eventBaseBlocks);

const revisedProjection = (comparison: FolioContentComparison): FolioContentBlock[] =>
  comparison.events.flatMap(eventRevisedBlocks);

const textBefore = (segments: readonly FolioContentTextSegment[]): string =>
  segments
    .filter(({ type }) => type !== "ins")
    .map(({ text }) => text)
    .join("");

const textAfter = (segments: readonly FolioContentTextSegment[]): string =>
  segments
    .filter(({ type }) => type !== "del")
    .map(({ text }) => text)
    .join("");

type SegmentOffsetExpectation = {
  segments: readonly FolioContentTextSegment[];
  base: string;
  revised: string;
};

const expectSegmentOffsets = ({ segments, base, revised }: SegmentOffsetExpectation): void => {
  let baseOffset = 0;
  let revisedOffset = 0;
  for (const segment of segments) {
    expect(segment.baseStart).toBe(baseOffset);
    expect(segment.revisedStart).toBe(revisedOffset);
    if (segment.type === "ins") {
      expect(segment.baseEnd).toBe(baseOffset);
    } else {
      expect(base.slice(segment.baseStart, segment.baseEnd)).toBe(segment.text);
      baseOffset = segment.baseEnd;
    }
    if (segment.type === "del") {
      expect(segment.revisedEnd).toBe(revisedOffset);
    } else {
      expect(revised.slice(segment.revisedStart, segment.revisedEnd)).toBe(segment.text);
      revisedOffset = segment.revisedEnd;
    }
  }
  expect(baseOffset).toBe(base.length);
  expect(revisedOffset).toBe(revised.length);
};

const expectRelationReconstructs = (relation: FolioContentPairRelation): void => {
  expect(textBefore(relation.segments)).toBe(
    relation.base.block.text.slice(relation.base.startOffset, relation.base.endOffset),
  );
  expect(textAfter(relation.segments)).toBe(
    relation.revised.block.text.slice(relation.revised.startOffset, relation.revised.endOffset),
  );
  let baseOffset = relation.base.startOffset;
  let revisedOffset = relation.revised.startOffset;
  for (const segment of relation.segments) {
    expect(segment.baseStart).toBe(baseOffset);
    expect(segment.revisedStart).toBe(revisedOffset);
    if (segment.type !== "ins") {
      expect(relation.base.block.text.slice(segment.baseStart, segment.baseEnd)).toBe(segment.text);
      baseOffset = segment.baseEnd;
    }
    if (segment.type !== "del") {
      expect(relation.revised.block.text.slice(segment.revisedStart, segment.revisedEnd)).toBe(
        segment.text,
      );
      revisedOffset = segment.revisedEnd;
    }
  }
  expect(baseOffset).toBe(relation.base.endOffset);
  expect(revisedOffset).toBe(relation.revised.endOffset);
};

const EXPECTED_EVENT_CARDINALITY = {
  unchanged: [1, 1],
  modified: [1, 1],
  formatting: [1, 1],
  inserted: [0, 1],
  deleted: [1, 0],
  movedFrom: [1, 0],
  movedTo: [0, 1],
  split: [1, 2],
  merge: [2, 1],
} as const satisfies Record<
  Exclude<FolioContentComparisonEvent["type"], "structural" | "tableReplacement">,
  readonly [number, number]
>;

const expectEventRelationsReconstruct = (event: FolioContentComparisonEvent): void => {
  switch (event.type) {
    case "unchanged":
    case "modified":
    case "formatting":
      expect(event.relation.relationType).toBe("whole");
      expectRelationReconstructs(event.relation);
      return;
    case "movedFrom":
    case "movedTo":
      expect(event.move.relation.relationType).toBe("whole");
      expectRelationReconstructs(event.move.relation);
      return;
    case "split":
    case "merge":
      expect(event.relations.every(({ relationType }) => relationType === "range")).toBe(true);
      expect(event.separator.relationType).toBe("separator");
      for (const relation of [...event.relations, event.separator]) {
        expectRelationReconstructs(relation);
      }
      return;
    case "inserted":
    case "deleted":
    case "structural":
      return;
    case "tableReplacement":
      for (const nested of event.replacement.refinement.events) {
        expectEventRelationsReconstruct(nested);
      }
      return;
    default: {
      const unreachable: never = event;
      throw new Error(`Unhandled event ${JSON.stringify(unreachable)}`);
    }
  }
};

const expectComparisonReconstructs = (
  comparison: FolioContentComparison,
  base: readonly FolioContentBlock[],
  revised: readonly FolioContentBlock[],
): void => {
  expect(baseProjection(comparison)).toEqual(base);
  expect(revisedProjection(comparison)).toEqual(revised);
  for (const event of comparison.events) {
    const cardinality = event.type === "tableReplacement"
      ? [event.replacement.baseBlocks.length, event.replacement.revisedBlocks.length]
      : event.type === "structural"
        ? event.change.type.endsWith("-insert")
          ? ([0, 1] as const)
          : ([1, 0] as const)
        : EXPECTED_EVENT_CARDINALITY[event.type];
    expect([eventBaseBlocks(event).length, eventRevisedBlocks(event).length]).toEqual(cardinality);
    expectEventRelationsReconstruct(event);
    if (event.type === "tableReplacement") {
      expectComparisonReconstructs(
        event.replacement.refinement,
        event.replacement.baseBlocks,
        event.replacement.revisedBlocks,
      );
    }
  }
};

const stableAnchors = (): { base: TestBlock[]; revised: TestBlock[] } => ({
  base: [
    contentBlock({ id: "anchor-a", text: "First durable anchor text" }),
    contentBlock({ id: "anchor-b", text: "Second durable anchor text" }),
    contentBlock({ id: "anchor-c", text: "Third durable anchor text" }),
  ],
  revised: [
    contentBlock({ id: "anchor-a", text: "First durable anchor text" }),
    contentBlock({ id: "anchor-b", text: "Second durable anchor text" }),
    contentBlock({ id: "anchor-c", text: "Third durable anchor text" }),
  ],
});

describe("representation-neutral comparison stream", () => {
  test("ignores consumer metadata outside canonical property sets", () => {
    const base = {
      ...contentBlock({ id: "clause", text: "Original clause" }),
      sourceAnchor: { id: "source-clause", ordinal: 4 },
    };

    const result = compareContent({
      base: { blocks: [base] },
      revised: { blocks: [] },
    });
    expect(result.isOk()).toBe(true);
    if (!result.isOk()) return;
    const event = result.value.events.at(0);
    expect(event?.type).toBe("deleted");
    if (event?.type !== "deleted") return;
    expect(Reflect.has(event.block, "sourceAnchor")).toBe(false);
  });

  test("equal content and presentation produce only unchanged events", () => {
    const base = [
      contentBlock({
        id: "heading",
        kind: "heading",
        text: "Terms",
        headingLevel: 1,
        styleId: "Heading1",
        runs: [{ text: "Terms", bold: true }],
      }),
      contentBlock({
        id: "body",
        text: "Payment is due.",
        directAlignment: "center",
        directSpacing: { spaceAfter: 120 },
      }),
    ];
    const revised = [
      contentBlock({
        id: "heading",
        kind: "heading",
        text: "Terms",
        headingLevel: 1,
        styleId: "Heading1",
        runs: [{ text: "Terms", bold: true }],
      }),
      contentBlock({
        id: "body",
        text: "Payment is due.",
        directAlignment: "center",
        directSpacing: { spaceAfter: 120 },
      }),
    ];

    const comparison = successfulComparison({ base, revised });

    expect(eventTypes(comparison)).toEqual(["unchanged", "unchanged"]);
    expect(baseProjection(comparison)).toEqual(base);
    expect(revisedProjection(comparison)).toEqual(revised);
    for (const event of comparison.events) {
      if (event.type !== "unchanged") throw new Error("expected unchanged events");
      expect(event.relation.relationType).toBe("whole");
      expectRelationReconstructs(event.relation);
    }
  });

  test("empty documents produce an empty ordered stream", () => {
    expect(successfulComparison({ base: [], revised: [] })).toEqual({ events: [] });
  });

  test("insertions and deletions stay in their logical stream positions", () => {
    const firstBase = contentBlock({ id: "first", text: "First clause" });
    const lastBase = contentBlock({ id: "last", text: "Last clause" });
    const firstRevised = contentBlock({ id: "first", text: "First clause" });
    const inserted = contentBlock({ id: "inserted", text: "New clause" });
    const lastRevised = contentBlock({ id: "last", text: "Last clause" });

    const insertion = successfulComparison({
      base: [firstBase, lastBase],
      revised: [firstRevised, inserted, lastRevised],
    });
    expect(eventTypes(insertion)).toEqual(["unchanged", "inserted", "unchanged"]);
    expect(baseProjection(insertion)).toEqual([firstBase, lastBase]);
    expect(revisedProjection(insertion)).toEqual([firstRevised, inserted, lastRevised]);

    const deletion = successfulComparison({
      base: [firstRevised, inserted, lastRevised],
      revised: [firstBase, lastBase],
    });
    expect(eventTypes(deletion)).toEqual(["unchanged", "deleted", "unchanged"]);
    expect(baseProjection(deletion)).toEqual([firstRevised, inserted, lastRevised]);
    expect(revisedProjection(deletion)).toEqual([firstBase, lastBase]);
  });

  test("a replacement carries ordered segments with offsets on both sides", () => {
    const base = contentBlock({ id: "term", text: "The fee is due." });
    const revised = contentBlock({ id: "term", text: "The tax is due." });

    const comparison = successfulComparison({ base: [base], revised: [revised] });

    expect(comparison.events).toEqual([
      {
        type: "modified",
        relation: {
          relationType: "whole",
          base: { block: base, startOffset: 0, endOffset: 15 },
          revised: { block: revised, startOffset: 0, endOffset: 15 },
          blockChanges: [],
          formatting: null,
          segments: [
            {
              type: "equal",
              text: "The",
              baseStart: 0,
              baseEnd: 3,
              revisedStart: 0,
              revisedEnd: 3,
            },
            {
              type: "del",
              text: " fee",
              baseStart: 3,
              baseEnd: 7,
              revisedStart: 3,
              revisedEnd: 3,
            },
            {
              type: "ins",
              text: " tax",
              baseStart: 7,
              baseEnd: 7,
              revisedStart: 3,
              revisedEnd: 7,
            },
            {
              type: "equal",
              text: " is due.",
              baseStart: 7,
              baseEnd: 15,
              revisedStart: 7,
              revisedEnd: 15,
            },
          ],
        },
      },
    ]);
  });

  test("paragraph splits and merges preserve their exact block tuples", () => {
    const joinedBase = contentBlock({ id: "first", text: "Alpha Beta" });
    const firstRevised = contentBlock({ id: "first", text: "Alpha" });
    const secondRevised = contentBlock({ id: "second", text: "Beta" });

    const split = successfulComparison({
      base: [joinedBase],
      revised: [firstRevised, secondRevised],
    });
    const splitEvent = split.events.at(0);
    expect(splitEvent?.type).toBe("split");
    if (splitEvent?.type !== "split") throw new Error("expected split");
    expect(baseProjection(split)).toEqual([joinedBase]);
    expect(revisedProjection(split)).toEqual([firstRevised, secondRevised]);
    expect(splitEvent.relations.map(({ relationType }) => relationType)).toEqual([
      "range",
      "range",
    ]);
    expect(splitEvent.separator.relationType).toBe("separator");
    expect(splitEvent.relations.map(({ base, revised }) => ({ base, revised }))).toEqual([
      {
        base: { block: joinedBase, startOffset: 0, endOffset: 5 },
        revised: { block: firstRevised, startOffset: 0, endOffset: 5 },
      },
      {
        base: { block: joinedBase, startOffset: 6, endOffset: 10 },
        revised: { block: secondRevised, startOffset: 0, endOffset: 4 },
      },
    ]);
    expect(splitEvent.separator).toMatchObject({
      base: { block: joinedBase, startOffset: 5, endOffset: 6 },
      revised: { block: firstRevised, startOffset: 5, endOffset: 5 },
      segments: [{ type: "del", text: " ", baseStart: 5, baseEnd: 6 }],
    });
    for (const relation of [...splitEvent.relations, splitEvent.separator]) {
      expectRelationReconstructs(relation);
    }

    const firstBase = contentBlock({ id: "first", text: "Alpha" });
    const secondBase = contentBlock({ id: "second", text: "Beta" });
    const joinedRevised = contentBlock({ id: "first", text: "Alpha Beta" });
    const merge = successfulComparison({
      base: [firstBase, secondBase],
      revised: [joinedRevised],
    });
    const mergeEvent = merge.events.at(0);
    expect(mergeEvent?.type).toBe("merge");
    if (mergeEvent?.type !== "merge") throw new Error("expected merge");
    expect(baseProjection(merge)).toEqual([firstBase, secondBase]);
    expect(revisedProjection(merge)).toEqual([joinedRevised]);
    expect(mergeEvent.relations.map(({ relationType }) => relationType)).toEqual([
      "range",
      "range",
    ]);
    expect(mergeEvent.separator.relationType).toBe("separator");
    expect(mergeEvent.relations.map(({ base, revised }) => ({ base, revised }))).toEqual([
      {
        base: { block: firstBase, startOffset: 0, endOffset: 5 },
        revised: { block: joinedRevised, startOffset: 0, endOffset: 5 },
      },
      {
        base: { block: secondBase, startOffset: 0, endOffset: 4 },
        revised: { block: joinedRevised, startOffset: 6, endOffset: 10 },
      },
    ]);
    expect(mergeEvent.separator).toMatchObject({
      base: { block: firstBase, startOffset: 5, endOffset: 5 },
      revised: { block: joinedRevised, startOffset: 5, endOffset: 6 },
      segments: [{ type: "ins", text: " ", revisedStart: 5, revisedEnd: 6 }],
    });
    for (const relation of [...mergeEvent.relations, mergeEvent.separator]) {
      expectRelationReconstructs(relation);
    }
  });

  test("split and merge classification is symmetric about the surviving block", () => {
    const joinedBase = contentBlock({ id: "right", text: "Alpha Beta" });
    const firstRevised = contentBlock({ id: "left", text: "Alpha" });
    const secondRevised = contentBlock({ id: "right", text: "Beta" });
    const split = successfulComparison({
      base: [joinedBase],
      revised: [firstRevised, secondRevised],
    });

    expect(split.events.map(({ type }) => type)).toEqual(["split"]);
    expect(baseProjection(split)).toEqual([joinedBase]);
    expect(revisedProjection(split)).toEqual([firstRevised, secondRevised]);

    const firstBase = contentBlock({ id: "left", text: "Alpha" });
    const secondBase = contentBlock({ id: "right", text: "Beta" });
    const joinedRevised = contentBlock({ id: "right", text: "Alpha Beta" });
    const merge = successfulComparison({
      base: [firstBase, secondBase],
      revised: [joinedRevised],
    });

    expect(merge.events.map(({ type }) => type)).toEqual(["merge"]);
    expect(baseProjection(merge)).toEqual([firstBase, secondBase]);
    expect(revisedProjection(merge)).toEqual([joinedRevised]);
  });

  test("split and merge segments use JavaScript UTF-16 offsets", () => {
    const joined = contentBlock({ id: "first", text: "A😀 B" });
    const first = contentBlock({ id: "first", text: "A😀" });
    const second = contentBlock({ id: "second", text: "B" });

    const split = successfulComparison({
      base: [joined],
      revised: [first, second],
      granularity: "character",
    }).events.at(0);
    expect(split?.type).toBe("split");
    if (split?.type !== "split") {
      throw new Error("expected a split event");
    }
    expect(split.relations[0].base.endOffset).toBe(3);
    for (const relation of [...split.relations, split.separator]) {
      expectRelationReconstructs(relation);
    }

    const merge = successfulComparison({
      base: [first, second],
      revised: [joined],
      granularity: "character",
    }).events.at(0);
    expect(merge?.type).toBe("merge");
    if (merge?.type !== "merge") {
      throw new Error("expected a merge event");
    }
    expect(merge.relations[1].revised.startOffset).toBe(4);
    for (const relation of [...merge.relations, merge.separator]) {
      expectRelationReconstructs(relation);
    }
  });

  test("the neutral core computes split and merge segments exactly once", () => {
    const fixtures = [
      {
        eventType: "split",
        base: [contentBlock({ id: "first", text: "A😀 B" })],
        revised: [
          contentBlock({ id: "first", text: "A😀" }),
          contentBlock({ id: "second", text: "B" }),
        ],
        expectedPairs: [
          ["A😀", "A😀"],
          ["B", "B"],
          [" ", ""],
        ],
      },
      {
        eventType: "merge",
        base: [
          contentBlock({ id: "first", text: "A😀" }),
          contentBlock({ id: "second", text: "B" }),
        ],
        revised: [contentBlock({ id: "first", text: "A😀 B" })],
        expectedPairs: [
          ["A😀", "A😀"],
          ["B", "B"],
          ["", " "],
        ],
      },
    ] as const;

    for (const fixture of fixtures) {
      const calls: [string, string][] = [];
      const workSession = createContentComparisonWorkSession({
        diffText: (base, revised) => {
          calls.push([base, revised]);
          return [
            ...(base.length > 0 ? [{ type: "del" as const, text: base }] : []),
            ...(revised.length > 0 ? [{ type: "ins" as const, text: revised }] : []),
          ];
        },
      });
      const operation = workSession.captureComparison({
        base: { blocks: fixture.base },
        revised: { blocks: fixture.revised },
      });
      if (operation.isErr()) {
        throw operation.error;
      }
      const comparison = operation.value.compare();
      if (comparison.isErr()) {
        throw comparison.error;
      }
      expect(calls, fixture.eventType).toEqual(fixture.expectedPairs);
      const event = comparison.value.events.at(0);
      expect(event?.type).toBe(fixture.eventType);
      if (event?.type !== "split" && event?.type !== "merge") {
        throw new Error(`expected a ${fixture.eventType} event`);
      }
      for (const [relation, expected] of [
        [event.relations[0], fixture.expectedPairs[0]],
        [event.relations[1], fixture.expectedPairs[1]],
        [event.separator, fixture.expectedPairs[2]],
      ] as const) {
        expect(
          relation.segments.map(({ type, text }) => ({ type, text })),
          fixture.eventType,
        ).toEqual([
          ...(expected[0].length > 0 ? [{ type: "del", text: expected[0] }] : []),
          ...(expected[1].length > 0 ? [{ type: "ins", text: expected[1] }] : []),
        ]);
      }
    }
  });

  test("split and merge events carry paragraph formatting for their paired blocks", () => {
    const baseFormatting = {
      styleId: "Base",
      listLevel: 0,
      directAlignment: "left",
      directSpacing: { spaceAfter: 120 },
    } as const;
    const joinedBase = contentBlock({
      id: "first",
      kind: "heading",
      headingLevel: 1,
      displayLabel: "1",
      text: "Alpha Beta",
      runs: [
        { text: "Alpha", bold: true },
        { text: " " },
        { text: "Beta", italic: true },
      ],
      ...baseFormatting,
    });
    const firstRevised = contentBlock({
      id: "first",
      text: "Alpha",
      runs: [{ text: "Alpha", italic: true }],
    });
    const secondRevised = contentBlock({
      id: "second",
      text: "Beta",
      runs: [{ text: "Beta", bold: true }],
      styleId: "Second",
      listLevel: 2,
      directAlignment: "right",
      directSpacing: { lineSpacing: 240, lineSpacingRule: "exact" },
    });

    const split = successfulComparison({
      base: [joinedBase],
      revised: [firstRevised, secondRevised],
    }).events.at(0);
    expect(split?.type).toBe("split");
    if (split?.type !== "split") {
      throw new Error("expected a split event");
    }
    expect(split.relations.map(changedBlockKeys)).toEqual([
      ["kind", "displayLabel", "headingLevel"],
      ["kind", "displayLabel", "headingLevel"],
    ]);
    expect(
      split.relations.map(({ formatting }) =>
        formatting?.paragraph.authored.map(({ key, revised }) => ({ key, revised })),
      ),
    ).toEqual([
      [
        { key: "directAlignment", revised: { type: "absent" } },
        { key: "directSpacing", revised: { type: "absent" } },
        { key: "listLevel", revised: { type: "absent" } },
        { key: "styleId", revised: { type: "absent" } },
      ],
      [
        { key: "directAlignment", revised: { type: "present", value: "right" } },
        {
          key: "directSpacing",
          revised: {
            type: "present",
            value: objectProperty({ lineSpacing: 240, lineSpacingRule: "exact" }),
          },
        },
        { key: "listLevel", revised: { type: "present", value: 2 } },
        { key: "styleId", revised: { type: "present", value: "Second" } },
      ],
    ]);
    expect(
      split.relations.map(({ formatting }) =>
        formatting?.ranges.map(({ baseStart, baseEnd, revisedStart, revisedEnd, formatting }) => ({
          baseStart,
          baseEnd,
          revisedStart,
          revisedEnd,
          effective: formatting.effective.map(({ key }) => key),
        })),
      ),
    ).toEqual([
      [{ baseStart: 0, baseEnd: 5, revisedStart: 0, revisedEnd: 5, effective: ["bold", "italic"] }],
      [{ baseStart: 6, baseEnd: 10, revisedStart: 0, revisedEnd: 4, effective: ["bold", "italic"] }],
    ]);

    const firstBase = contentBlock({
      id: "first",
      kind: "heading",
      headingLevel: 1,
      displayLabel: "1",
      text: "Alpha",
      runs: [{ text: "Alpha", bold: true }],
      ...baseFormatting,
    });
    const secondBase = contentBlock({
      id: "second",
      kind: "heading",
      headingLevel: 2,
      displayLabel: "1.1",
      text: "Beta",
      runs: [{ text: "Beta", italic: true }],
      styleId: "Second",
    });
    const joinedRevised = contentBlock({
      id: "first",
      text: "Alpha Beta",
      runs: [
        { text: "Alpha", italic: true },
        { text: " " },
        { text: "Beta", bold: true },
      ],
    });
    const merge = successfulComparison({
      base: [firstBase, secondBase],
      revised: [joinedRevised],
    }).events.at(0);
    expect(merge?.type).toBe("merge");
    if (merge?.type !== "merge") {
      throw new Error("expected a merge event");
    }
    expect(merge.relations.map(changedBlockKeys)).toEqual([
      ["kind", "displayLabel", "headingLevel"],
      ["kind", "displayLabel", "headingLevel"],
    ]);
    expect(
      merge.relations.map(({ formatting }) =>
        formatting?.paragraph.authored.map(({ key, revised }) => ({ key, revised })),
      ),
    ).toEqual([
      [
        { key: "directAlignment", revised: { type: "absent" } },
        { key: "directSpacing", revised: { type: "absent" } },
        { key: "listLevel", revised: { type: "absent" } },
        { key: "styleId", revised: { type: "absent" } },
      ],
      [{ key: "styleId", revised: { type: "absent" } }],
    ]);
    expect(
      merge.relations.map(({ formatting }) =>
        formatting?.ranges.map(({ baseStart, baseEnd, revisedStart, revisedEnd, formatting }) => ({
          baseStart,
          baseEnd,
          revisedStart,
          revisedEnd,
          effective: formatting.effective.map(({ key }) => key),
        })),
      ),
    ).toEqual([
      [{ baseStart: 0, baseEnd: 5, revisedStart: 0, revisedEnd: 5, effective: ["bold", "italic"] }],
      [{ baseStart: 0, baseEnd: 4, revisedStart: 6, revisedEnd: 10, effective: ["bold", "italic"] }],
    ]);
  });

  test("exact and edited moves share one identity between their two stream positions", () => {
    const exactBase = contentBlock({
      id: "exact-base",
      identitySemantics: "persistent-hint",
      text: "This clause remains exactly here",
    });
    const editedBase = contentBlock({
      id: "edited-base",
      identitySemantics: "persistent-hint",
      text: "alpha beta gamma delta epsilon",
    });
    const anchors = stableAnchors();
    const exactRevised = contentBlock({
      id: "exact-revised",
      identitySemantics: "persistent-hint",
      text: "This clause remains exactly here",
    });
    const editedRevised = contentBlock({
      id: "edited-revised",
      identitySemantics: "persistent-hint",
      text: "alpha beta gamma delta zeta",
    });

    const comparison = successfulComparison({
      base: [exactBase, editedBase, ...anchors.base],
      revised: [...anchors.revised, exactRevised, editedRevised],
    });

    expect(eventTypes(comparison)).toEqual([
      "movedFrom",
      "movedFrom",
      "unchanged",
      "unchanged",
      "unchanged",
      "movedTo",
      "movedTo",
    ]);
    const movedFrom = comparison.events.filter(({ type }) => type === "movedFrom");
    const movedTo = comparison.events.filter(({ type }) => type === "movedTo");
    expect(
      movedFrom.map(({ move }) => ({
        moveId: move.id,
        baseBlockId: move.relation.base.block.identity.id,
      })),
    ).toEqual([
      { moveId: 1, baseBlockId: "exact-base" },
      { moveId: 2, baseBlockId: "edited-base" },
    ]);
    expect(
      movedTo.map(({ move }) => ({
        moveId: move.id,
        baseBlockId: move.relation.base.block.identity.id,
        revisedBlockId: move.relation.revised.block.identity.id,
      })),
    ).toEqual([
      { moveId: 1, baseBlockId: "exact-base", revisedBlockId: "exact-revised" },
      { moveId: 2, baseBlockId: "edited-base", revisedBlockId: "edited-revised" },
    ]);
    expect(movedTo[0]?.move.relation.segments).toEqual([
      {
        type: "equal",
        text: "This clause remains exactly here",
        baseStart: 0,
        baseEnd: 32,
        revisedStart: 0,
        revisedEnd: 32,
      },
    ]);
    expect(movedTo[1]?.move.relation.segments).toEqual([
      {
        type: "equal",
        text: "alpha beta gamma delta",
        baseStart: 0,
        baseEnd: 22,
        revisedStart: 0,
        revisedEnd: 22,
      },
      {
        type: "del",
        text: " epsilon",
        baseStart: 22,
        baseEnd: 30,
        revisedStart: 22,
        revisedEnd: 22,
      },
      {
        type: "ins",
        text: " zeta",
        baseStart: 30,
        baseEnd: 30,
        revisedStart: 22,
        revisedEnd: 27,
      },
    ]);
  });

  test("stable IDs classify short exact and edited relocations without a text heuristic", () => {
    const exactBase = contentBlock({ id: "short-exact", text: "Title" });
    const editedBase = contentBlock({ id: "short-edited", text: "Old" });
    const anchors = stableAnchors();
    const exactRevised = contentBlock({ id: "short-exact", text: "Title" });
    const editedRevised = contentBlock({ id: "short-edited", text: "New" });

    const comparison = successfulComparison({
      base: [exactBase, editedBase, ...anchors.base],
      revised: [...anchors.revised, exactRevised, editedRevised],
      granularity: "character",
    });
    const movedFrom = comparison.events.filter(({ type }) => type === "movedFrom");
    const movedTo = comparison.events.filter(({ type }) => type === "movedTo");

    expect(eventTypes(comparison)).toEqual([
      "movedFrom",
      "movedFrom",
      "unchanged",
      "unchanged",
      "unchanged",
      "movedTo",
      "movedTo",
    ]);
    expect(
      movedFrom.map(({ move }) => [move.id, move.relation.base.block.identity.id]),
    ).toEqual([
      [1, "short-exact"],
      [2, "short-edited"],
    ]);
    expect(
      movedTo.map(({ move }) => [
        move.id,
        move.relation.base.block.identity.id,
        move.relation.revised.block.identity.id,
      ]),
    ).toEqual([
      [1, "short-exact", "short-exact"],
      [2, "short-edited", "short-edited"],
    ]);
    expect(movedTo[0]?.move.relation.segments.every(({ type }) => type === "equal")).toBe(true);
    const editedSegments = movedTo[1]?.move.relation.segments;
    expect(editedSegments).toBeDefined();
    if (editedSegments) {
      expect(textBefore(editedSegments)).toBe("Old");
      expect(textAfter(editedSegments)).toBe("New");
      expectSegmentOffsets({ segments: editedSegments, base: "Old", revised: "New" });
    }
  });

  test("stable move identity takes precedence over an earlier exact-text candidate", () => {
    const source = contentBlock({
      id: "stable-source",
      text: "alpha beta gamma delta epsilon",
    });
    const decoy = contentBlock({
      id: "exact-decoy",
      text: "alpha beta gamma delta epsilon",
    });
    const stableTarget = contentBlock({
      id: "stable-source",
      text: "alpha beta gamma delta zeta",
    });
    const anchors = stableAnchors();

    const comparison = successfulComparison({
      base: [source, ...anchors.base],
      revised: [...anchors.revised, decoy, stableTarget],
    });
    const movedTo = comparison.events.find(({ type }) => type === "movedTo");

    expect(movedTo).toMatchObject({
      type: "movedTo",
      move: {
        relation: {
          base: { block: { identity: { id: "stable-source" } } },
          revised: { block: { identity: { id: "stable-source" } } },
        },
      },
    });
    expect(
      comparison.events.find(
        (event) => event.type === "inserted" && event.block.identity.id === "exact-decoy",
      ),
    ).toBeDefined();
  });

  test("an exact move takes precedence over an earlier edited candidate", () => {
    const source = contentBlock({
      id: "exact-source",
      identitySemantics: "persistent-hint",
      text: "alpha beta gamma delta epsilon",
    });
    const editedDecoy = contentBlock({
      id: "edited-decoy",
      identitySemantics: "persistent-hint",
      text: "alpha beta gamma delta zeta",
    });
    const exactTarget = contentBlock({
      id: "exact-target",
      identitySemantics: "persistent-hint",
      text: "alpha beta gamma delta epsilon",
    });
    const anchors = stableAnchors();

    const comparison = successfulComparison({
      base: [source, ...anchors.base],
      revised: [...anchors.revised, editedDecoy, exactTarget],
    });
    const movedTo = comparison.events.find(({ type }) => type === "movedTo");

    expect(movedTo).toMatchObject({
      type: "movedTo",
      move: {
        relation: {
          base: { block: { identity: { id: "exact-source" } } },
          revised: { block: { identity: { id: "exact-target" } } },
        },
      },
    });
    expect(
      comparison.events.find(
        (event) => event.type === "inserted" && event.block.identity.id === "edited-decoy",
      ),
    ).toBeDefined();
  });

  test("a positional short ID match remains a deletion and insertion", () => {
    const movedBase = contentBlock({
      id: "position-0",
      text: "Title",
      identitySemantics: "positional",
    });
    const movedRevised = contentBlock({
      id: "position-0",
      text: "Title",
      identitySemantics: "positional",
    });
    const anchors = stableAnchors();

    const comparison = successfulComparison({
      base: [movedBase, ...anchors.base],
      revised: [...anchors.revised, movedRevised],
    });

    expect(eventTypes(comparison)).toEqual([
      "deleted",
      "unchanged",
      "unchanged",
      "unchanged",
      "inserted",
    ]);
  });

  test("repeated move candidates pair FIFO", () => {
    const repeatedText = "standard terms apply equally here";
    const firstBase = contentBlock({
      id: "first-base",
      identitySemantics: "persistent-hint",
      text: repeatedText,
    });
    const secondBase = contentBlock({
      id: "second-base",
      identitySemantics: "persistent-hint",
      text: repeatedText,
    });
    const firstRevised = contentBlock({
      id: "first-revised",
      identitySemantics: "persistent-hint",
      text: repeatedText,
    });
    const secondRevised = contentBlock({
      id: "second-revised",
      identitySemantics: "persistent-hint",
      text: repeatedText,
    });
    const anchors = stableAnchors();

    const comparison = successfulComparison({
      base: [firstBase, secondBase, ...anchors.base],
      revised: [...anchors.revised, firstRevised, secondRevised],
    });
    const movedTo = comparison.events.filter(({ type }) => type === "movedTo");

    expect(movedTo.map(({ move }) => move.relation.base.block.identity.id)).toEqual([
      "first-base",
      "second-base",
    ]);
    expect(movedTo.map(({ move }) => move.relation.revised.block.identity.id)).toEqual([
      "first-revised",
      "second-revised",
    ]);
  });

  test("a stable move source is reserved for its stable revised block", () => {
    const movedText = "This clause has enough words";
    const movedBase = contentBlock({ id: "stable-move", text: movedText });
    const anchors = stableAnchors();
    const positionalCopy = contentBlock({
      id: "positional-copy",
      identitySemantics: "positional",
      text: movedText,
    });
    const movedRevised = contentBlock({ id: "stable-move", text: movedText });

    const comparison = successfulComparison({
      base: [movedBase, ...anchors.base],
      revised: [...anchors.revised, positionalCopy, movedRevised],
    });

    expect(eventTypes(comparison)).toEqual([
      "movedFrom",
      "unchanged",
      "unchanged",
      "unchanged",
      "inserted",
      "movedTo",
    ]);
    const movedTo = comparison.events.find(({ type }) => type === "movedTo");
    expect(movedTo).toMatchObject({
      move: {
        id: 1,
        relation: {
          base: { block: movedBase },
          revised: { block: movedRevised },
        },
      },
    });
  });

  test("an exact move source is reserved from an earlier edited candidate", () => {
    const exactText = "alpha beta gamma delta epsilon";
    const movedBase = contentBlock({
      id: "source",
      identitySemantics: "positional",
      text: exactText,
    });
    const anchors = stableAnchors();
    const editedCandidate = contentBlock({
      id: "edited",
      identitySemantics: "positional",
      text: "alpha beta gamma delta zeta",
    });
    const exactRevised = contentBlock({
      id: "exact",
      identitySemantics: "positional",
      text: exactText,
    });

    const comparison = successfulComparison({
      base: [movedBase, ...anchors.base],
      revised: [...anchors.revised, editedCandidate, exactRevised],
    });

    expect(eventTypes(comparison)).toEqual([
      "movedFrom",
      "unchanged",
      "unchanged",
      "unchanged",
      "inserted",
      "movedTo",
    ]);
    const movedTo = comparison.events.find(({ type }) => type === "movedTo");
    expect(movedTo).toMatchObject({
      move: {
        id: 1,
        relation: {
          base: { block: movedBase },
          revised: { block: exactRevised },
        },
      },
    });
  });

  test("short relocated boilerplate is not reported as a move", () => {
    const shortBase = contentBlock({ id: "short-base", text: "standard terms" });
    const shortRevised = contentBlock({ id: "short-revised", text: "standard terms" });
    const anchors = stableAnchors();

    const comparison = successfulComparison({
      base: [shortBase, ...anchors.base],
      revised: [...anchors.revised, shortRevised],
    });

    expect(eventTypes(comparison)).toEqual([
      "deleted",
      "unchanged",
      "unchanged",
      "unchanged",
      "inserted",
    ]);
  });

  test("paragraph-only formatting changes carry the revised property values", () => {
    const base = contentBlock({
      id: "clause",
      text: "Payment is due.",
      styleId: "Body",
      listLevel: 0,
      directAlignment: "left",
      directSpacing: { spaceBefore: 0, lineSpacingRule: "auto" },
    });
    const revised = contentBlock({
      id: "clause",
      text: "Payment is due.",
      styleId: "Clause",
      listLevel: 1,
      directAlignment: "center",
      directSpacing: { spaceAfter: 120, lineSpacingRule: "exact" },
    });

    const comparison = successfulComparison({ base: [base], revised: [revised] });

    const event = comparison.events.at(0);
    expect(event?.type).toBe("formatting");
    if (event?.type !== "formatting") throw new Error("expected formatting");
    expect(
      event.relation.formatting?.paragraph.authored.map(({ key, revised }) => ({ key, revised })),
    ).toEqual([
      { key: "directAlignment", revised: { type: "present", value: "center" } },
      {
        key: "directSpacing",
        revised: {
          type: "present",
          value: objectProperty({ spaceAfter: 120, lineSpacingRule: "exact" }),
        },
      },
      { key: "listLevel", revised: { type: "present", value: 1 } },
      { key: "styleId", revised: { type: "present", value: "Clause" } },
    ]);
    expect(event.relation.formatting?.ranges).toEqual([]);
  });

  test("authored and effective paragraph changes remain separate", () => {
    const base = contentBlock({
      id: "clause",
      text: "Payment is due.",
      directAlignment: "left",
      effectiveParagraphFormatting: { alignment: "left", fontSize: 11 },
    });
    const revised = contentBlock({
      id: "clause",
      text: "Payment is due.",
      directAlignment: "center",
      effectiveParagraphFormatting: { alignment: "center", fontSize: 12 },
    });

    const comparison = successfulComparison({ base: [base], revised: [revised] });
    const event = comparison.events.at(0);
    expect(event?.type).toBe("formatting");
    if (event?.type !== "formatting" || event.relation.formatting === null) return;
    expect(event.relation.formatting.paragraph.authored.map(({ key }) => key)).toEqual([
      "directAlignment",
    ]);
    expect(event.relation.formatting.paragraph.effective.map(({ key }) => key)).toEqual([
      "alignment",
      "fontSize",
    ]);
  });

  test("inline-only formatting changes carry UTF-16 range offsets", () => {
    const base = contentBlock({
      id: "clause",
      text: "A😀B",
      runs: [{ text: "A" }, { text: "😀B", bold: true }],
    });
    const revised = contentBlock({
      id: "clause",
      text: "A😀B",
      runs: [{ text: "A" }, { text: "😀B", italic: true }],
    });

    const comparison = successfulComparison({ base: [base], revised: [revised] });

    const event = comparison.events.at(0);
    expect(event?.type).toBe("formatting");
    if (event?.type !== "formatting") throw new Error("expected formatting");
    expect(event.relation.formatting).toEqual({
      paragraph: { authored: [], effective: [] },
      ranges: [
        {
          baseStart: 1,
          baseEnd: 4,
          revisedStart: 1,
          revisedEnd: 4,
          formatting: {
            authored: [],
            effective: [
              {
                key: "bold",
                base: { type: "present", value: true },
                revised: { type: "absent" },
              },
              {
                key: "italic",
                base: { type: "absent" },
                revised: { type: "present", value: true },
              },
            ],
          },
        },
      ],
    });
  });

  test("non-hex color tokens remain representation-neutral formatting values", () => {
    const base = contentBlock({
      id: "clause",
      text: "Payment",
      runs: [{ text: "Payment", color: "red" }],
    });
    const revised = contentBlock({
      id: "clause",
      text: "Payment",
      runs: [{ text: "Payment", color: "blue" }],
    });

    const comparison = successfulComparison({ base: [base], revised: [revised] });

    const event = comparison.events.at(0);
    expect(event?.type).toBe("formatting");
    if (event?.type !== "formatting") throw new Error("expected formatting");
    expect(event.relation.formatting?.ranges).toEqual([
      {
        baseStart: 0,
        baseEnd: 7,
        revisedStart: 0,
        revisedEnd: 7,
        formatting: {
          authored: [],
          effective: [
            {
              key: "color",
              base: { type: "present", value: "red" },
              revised: { type: "present", value: "blue" },
            },
          ],
        },
      },
    ]);
  });

  test("an explicit direct-color removal remains distinct from an absent property", () => {
    const base = contentBlock({
      id: "clause",
      text: "Payment",
      runs: [{ text: "Payment", color: "red", directFormatting: {} }],
    });
    const revised = contentBlock({
      id: "clause",
      text: "Payment",
      runs: [{ text: "Payment", color: "red", directFormatting: { color: null } }],
    });

    const comparison = successfulComparison({ base: [base], revised: [revised] });

    const event = comparison.events.at(0);
    expect(event?.type).toBe("formatting");
    if (event?.type !== "formatting") throw new Error("expected formatting");
    expect(event.relation.formatting?.ranges).toEqual([
      {
        baseStart: 0,
        baseEnd: 7,
        revisedStart: 0,
        revisedEnd: 7,
        formatting: {
          authored: [
            {
              key: "color",
              base: { type: "absent" },
              revised: { type: "present", value: null },
            },
          ],
          effective: [],
        },
      },
    ]);
  });

  test("text segment offsets use UTF-16 boundaries compatible with string slicing", () => {
    const base = contentBlock({ id: "unicode", text: "A😀B" });
    const revised = contentBlock({ id: "unicode", text: "A😀XB" });

    const comparison = successfulComparison({
      base: [base],
      revised: [revised],
      granularity: "character",
    });
    const event = comparison.events[0];

    expect(event?.type).toBe("modified");
    if (event?.type !== "modified") {
      return;
    }
    expect(event.relation.segments).toEqual([
      {
        type: "equal",
        text: "A😀",
        baseStart: 0,
        baseEnd: 3,
        revisedStart: 0,
        revisedEnd: 3,
      },
      {
        type: "ins",
        text: "X",
        baseStart: 3,
        baseEnd: 3,
        revisedStart: 3,
        revisedEnd: 4,
      },
      {
        type: "equal",
        text: "B",
        baseStart: 3,
        baseEnd: 4,
        revisedStart: 4,
        revisedEnd: 5,
      },
    ]);
    expectSegmentOffsets({
      segments: event.relation.segments,
      base: base.text,
      revised: revised.text,
    });
  });
});

describe("container-aware comparison", () => {
  test("table-cell coordinates remain authoritative when blocks share outer ancestry", () => {
    const containerPath = [{ kind: "section", id: "schedule" }] as const;
    const base = tableBlock({
      id: "joined",
      text: "Alpha Beta",
      rowIndex: 0,
      cellIndex: 0,
      containerPath,
    });
    const firstRevised = tableBlock({
      id: "joined",
      text: "Alpha",
      rowIndex: 0,
      cellIndex: 0,
      containerPath,
    });
    const secondRevised = tableBlock({
      id: "separate-cell",
      text: "Beta",
      rowIndex: 0,
      cellIndex: 1,
      containerPath,
    });

    const comparison = successfulComparison({
      base: [base],
      revised: [firstRevised, secondRevised],
    });

    expect(eventTypes(comparison)).not.toContain("split");
    expect(baseProjection(comparison)).toEqual([base]);
    expect(revisedProjection(comparison)).toEqual([firstRevised, secondRevised]);
  });

  test("an all-table snapshot preserves an unrepresentable span change", () => {
    const base = tableBlock({
      id: "base-cell",
      text: "Clause",
      rowIndex: 0,
      cellIndex: 0,
    });
    const revised = tableBlock({
      id: "revised-cell",
      text: "Clause",
      rowIndex: 0,
      cellIndex: 0,
      columnSpan: 2,
    });

    const comparison = successfulComparison({ base: [base], revised: [revised] });

    expect(eventTypes(comparison)).toEqual(["tableReplacement"]);
    const event = comparison.events.at(0);
    if (event?.type !== "tableReplacement") {
      throw new Error("Expected the incompatible table pair to own its refinement");
    }
    expect(eventTypes(event.replacement.refinement)).toEqual(["inserted", "deleted"]);
    expect(baseProjection(comparison)).toEqual([base]);
    expect(revisedProjection(comparison)).toEqual([revised]);
  });

  test("stable identities never pair blocks across table cells", () => {
    const base = [
      tableBlock({ id: "left", text: "Alpha", rowIndex: 0, cellIndex: 0 }),
      tableBlock({ id: "right", text: "Beta", rowIndex: 0, cellIndex: 1 }),
    ];
    const revised = [
      tableBlock({ id: "right", text: "Beta", rowIndex: 0, cellIndex: 0 }),
      tableBlock({ id: "left", text: "Alpha", rowIndex: 0, cellIndex: 1 }),
    ];

    const comparison = successfulComparison({ base, revised });

    expect(eventTypes(comparison)).toEqual(["inserted", "inserted", "deleted", "deleted"]);
    expect(baseProjection(comparison)).toEqual(base);
    expect(revisedProjection(comparison)).toEqual(revised);
  });

  test("authoritative table identities preserve shifted tables without cross-pairing", () => {
    const identifiedTableBlock = (id: string, tableIndex: number): TestBlock =>
      tableBlock({
        id,
        text: "Repeated clause",
        rowIndex: 0,
        cellIndex: 0,
        tableIndex,
        outerTableIdentity: { type: "authoritative", id: `outer-${id}` },
        tableIdentity: { type: "authoritative", id: `table-${id}` },
        rowIdentity: { type: "authoritative", id: `row-${id}` },
        cellIdentity: { type: "authoritative", id: `cell-${id}` },
      });
    const base = [identifiedTableBlock("a", 0), identifiedTableBlock("b", 1)];
    const revised = [
      identifiedTableBlock("new", 0),
      identifiedTableBlock("a", 1),
      identifiedTableBlock("b", 2),
    ];

    const comparison = successfulComparison({ base, revised });

    expect(baseProjection(comparison)).toEqual(base);
    expect(revisedProjection(comparison)).toEqual(revised);
    const pairedIds = comparison.events.flatMap((event) =>
      event.type === "modified" || event.type === "unchanged" || event.type === "formatting"
        ? [
            [
              event.relation.base.block.identity.id,
              event.relation.revised.block.identity.id,
            ],
          ]
        : [],
    );
    expect(pairedIds).toEqual([
      ["a", "a"],
      ["b", "b"],
    ]);
  });

  test("repeated table text with conflicting authoritative ancestry stays unpaired", () => {
    const nestedTableBlock = (
      id: string,
      tableIndex: number,
      containerId: string,
    ): TestBlock => {
      return tableBlock({
        id,
        text: "Repeated clause",
        rowIndex: 0,
        cellIndex: 0,
        tableIndex,
        containerPath: [{ kind: "section", id: containerId }],
      });
    };
    const base = [nestedTableBlock("a", 0, "base-a"), nestedTableBlock("b", 1, "base-b")];
    const revised = [
      nestedTableBlock("c", 0, "revised-c"),
      nestedTableBlock("d", 1, "revised-d"),
    ];

    const comparison = successfulComparison({ base, revised });

    expect(
      comparison.events.some(
        ({ type }) => type === "unchanged" || type === "modified" || type === "formatting",
      ),
    ).toBe(false);
    expect(baseProjection(comparison)).toEqual(base);
    expect(revisedProjection(comparison)).toEqual(revised);
  });

  test.each([
    {
      label: "stable-id",
      baseId: "relocated",
      revisedId: "relocated",
      baseText: "Title",
      revisedText: "Updated",
    },
    {
      label: "exact-text",
      baseId: "exact-base",
      revisedId: "exact-revised",
      baseText: "payment is due within thirty days",
      revisedText: "payment is due within thirty days",
    },
    {
      label: "edited-text",
      baseId: "edited-base",
      revisedId: "edited-revised",
      baseText: "payment is due within thirty days",
      revisedText: "payment is due within forty days",
    },
  ])(
    "keeps $label relocation across table cells separate",
    ({ baseId, revisedId, baseText, revisedText }) => {
      const base = [
        tableBlock({
          id: baseId,
          text: baseText,
          rowIndex: 0,
          cellIndex: 0,
          paragraphIndex: 0,
        }),
        tableBlock({
          id: "left-anchor",
          text: "Left cell durable anchor",
          rowIndex: 0,
          cellIndex: 0,
          paragraphIndex: 1,
        }),
        tableBlock({
          id: "right-anchor",
          text: "Right cell durable anchor",
          rowIndex: 0,
          cellIndex: 1,
          paragraphIndex: 0,
        }),
      ];
      const revised = [
        tableBlock({
          id: "left-anchor",
          text: "Left cell durable anchor",
          rowIndex: 0,
          cellIndex: 0,
          paragraphIndex: 0,
        }),
        tableBlock({
          id: "right-anchor",
          text: "Right cell durable anchor",
          rowIndex: 0,
          cellIndex: 1,
          paragraphIndex: 0,
        }),
        tableBlock({
          id: revisedId,
          text: revisedText,
          rowIndex: 0,
          cellIndex: 1,
          paragraphIndex: 1,
        }),
      ];

      const comparison = successfulComparison({ base, revised });

      expect(eventTypes(comparison)).toEqual(["deleted", "modified", "unchanged", "inserted"]);
      expect(baseProjection(comparison)).toEqual(base);
      expect(revisedProjection(comparison)).toEqual(revised);
    },
  );

  test("an inserted table row is one structural change with row-major events", () => {
    const base = [
      tableBlock({ id: "header-a", text: "Header A", rowIndex: 0, cellIndex: 0 }),
      tableBlock({ id: "header-b", text: "Header B", rowIndex: 0, cellIndex: 1 }),
      tableBlock({ id: "existing-a", text: "Existing A", rowIndex: 1, cellIndex: 0 }),
      tableBlock({ id: "existing-b", text: "Existing B", rowIndex: 1, cellIndex: 1 }),
    ];
    const insertedA = tableBlock({
      id: "inserted-a",
      text: "Added X",
      rowIndex: 1,
      cellIndex: 0,
    });
    const insertedB = tableBlock({
      id: "inserted-b",
      text: "Added Y",
      rowIndex: 1,
      cellIndex: 1,
    });
    const revised = [
      tableBlock({ id: "header-a", text: "Header A", rowIndex: 0, cellIndex: 0 }),
      tableBlock({ id: "header-b", text: "Header B", rowIndex: 0, cellIndex: 1 }),
      insertedA,
      insertedB,
      tableBlock({ id: "existing-a", text: "Existing A", rowIndex: 2, cellIndex: 0 }),
      tableBlock({ id: "existing-b", text: "Existing B", rowIndex: 2, cellIndex: 1 }),
    ];

    const comparison = successfulComparison({ base, revised });

    expect(eventTypes(comparison)).toEqual([
      "unchanged",
      "unchanged",
      "structural",
      "structural",
      "modified",
      "modified",
    ]);
    expect(revisedProjection(comparison).map(({ id }) => id)).toEqual(revised.map(({ id }) => id));
    const rowEvents = comparison.events.filter(({ type }) => type === "structural");
    expect(rowEvents.map(({ change }) => change)).toEqual([rowEvents[0]?.change, rowEvents[0]?.change]);
    expect(rowEvents[0]?.change).toMatchObject({
      type: "table-row-insert",
      tableIndex: 0,
      rowIndex: 1,
      blocks: [insertedA, insertedB],
    });
  });

  test("an inserted table column stays grouped while events remain row-major", () => {
    const base = [
      tableBlock({ id: "a0", text: "A0", rowIndex: 0, cellIndex: 0 }),
      tableBlock({ id: "a1", text: "A1", rowIndex: 0, cellIndex: 1 }),
      tableBlock({ id: "b0", text: "B0", rowIndex: 1, cellIndex: 0 }),
      tableBlock({ id: "b1", text: "B1", rowIndex: 1, cellIndex: 1 }),
    ];
    const insertedA = tableBlock({
      id: "ax",
      text: "AX",
      rowIndex: 0,
      cellIndex: 1,
      gridColumnIndex: 1,
    });
    const insertedB = tableBlock({
      id: "bx",
      text: "BX",
      rowIndex: 1,
      cellIndex: 1,
      gridColumnIndex: 1,
    });
    const revised = [
      tableBlock({ id: "a0", text: "A0", rowIndex: 0, cellIndex: 0 }),
      insertedA,
      tableBlock({
        id: "a1",
        text: "A1",
        rowIndex: 0,
        cellIndex: 2,
        gridColumnIndex: 2,
      }),
      tableBlock({ id: "b0", text: "B0", rowIndex: 1, cellIndex: 0 }),
      insertedB,
      tableBlock({
        id: "b1",
        text: "B1",
        rowIndex: 1,
        cellIndex: 2,
        gridColumnIndex: 2,
      }),
    ];

    const comparison = successfulComparison({ base, revised });

    expect(eventTypes(comparison)).toEqual([
      "unchanged",
      "structural",
      "modified",
      "unchanged",
      "structural",
      "modified",
    ]);
    expect(revisedProjection(comparison)).toEqual(revised);
    expect(baseProjection(comparison)).toEqual(base);
    const columnEvents = comparison.events.filter(({ type }) => type === "structural");
    expect(columnEvents.map(({ change }) => change)).toEqual([
      columnEvents[0]?.change,
      columnEvents[0]?.change,
    ]);
    expect(columnEvents[0]?.change).toMatchObject({
      type: "table-column-insert",
      tableIndex: 0,
      columnIndex: 1,
      blocks: [insertedA, insertedB],
    });
  });

  test("stable identity does not turn a table insertion into a cross-structure move", () => {
    const base = contentBlock({ id: "shared", text: "Clause" });
    const revised = tableBlock({
      id: "shared",
      text: "Clause",
      rowIndex: 0,
      cellIndex: 0,
    });

    const comparison = successfulComparison({ base: [base], revised: [revised] });

    expect(eventTypes(comparison)).toEqual(["structural", "deleted"]);
    const structural = comparison.events.at(0);
    expect(structural).toMatchObject({
      type: "structural",
      memberIndex: 0,
      change: {
        type: "table-insert",
        tableIndex: 0,
        blocks: [revised],
      },
    });
    expect(baseProjection(comparison)).toEqual([base]);
    expect(revisedProjection(comparison)).toEqual([revised]);
  });

  test("malformed table coordinate order returns an input error instead of panicking", () => {
    const base = [
      tableBlock({ id: "right", text: "Right", rowIndex: 0, cellIndex: 1 }),
      tableBlock({ id: "left", text: "Left", rowIndex: 0, cellIndex: 0 }),
    ];
    const revised = [
      tableBlock({ id: "left", text: "Left", rowIndex: 0, cellIndex: 0 }),
      tableBlock({ id: "right", text: "Right", rowIndex: 0, cellIndex: 1 }),
    ];

    const result = compareContent({ base: { blocks: base }, revised: { blocks: revised } });

    expect(result.isErr()).toBe(true);
    if (!result.isErr()) {
      return;
    }
    expect(result.error).toBeInstanceOf(InvalidFolioContentComparisonError);
    expect(result.error).toMatchObject({
      input: "base",
      blockIndex: 0,
      field: "blocks[0].table",
    });
  });

  test("physical table-cell order must agree with logical grid placement", () => {
    const result = compareContent({
      base: {
        blocks: [
          tableBlock({
            id: "physical-first",
            text: "First",
            rowIndex: 0,
            cellIndex: 0,
            gridColumnIndex: 1,
          }),
          tableBlock({
            id: "physical-second",
            text: "Second",
            rowIndex: 0,
            cellIndex: 1,
            gridColumnIndex: 0,
          }),
        ],
      },
      revised: { blocks: [] },
    });

    expect(result.isErr()).toBe(true);
    if (!result.isErr()) return;
    expect(result.error).toBeInstanceOf(InvalidFolioContentComparisonError);
    expect(result.error).toMatchObject({
      input: "base",
      blockIndex: 1,
      field: "blocks[1].table",
    });
  });

  test.each([
    { field: "gridColumnIndex", value: 1 },
    { field: "columnSpan", value: 2 },
    { field: "rowSpan", value: 2 },
  ] as const)("rejects inconsistent $field values within one physical cell", ({ field, value }) => {
    const first = tableBlock({
      id: "first",
      text: "First",
      rowIndex: 0,
      cellIndex: 0,
      paragraphIndex: 0,
    });
    const second = tableBlock({
      id: "second",
      text: "Second",
      rowIndex: 0,
      cellIndex: 0,
      paragraphIndex: 1,
      [field]: value,
    });

    const result = compareContent({
      base: { blocks: [first, second] },
      revised: { blocks: [] },
    });

    expect(result.isErr()).toBe(true);
    if (!result.isErr()) {
      return;
    }
    expect(result.error).toBeInstanceOf(InvalidFolioContentComparisonError);
    expect(result.error).toMatchObject({
      input: "base",
      blockIndex: 1,
      field: "blocks[1].table",
    });
  });

  test("a stable block relocated across generic containers is an explicit move", () => {
    const base = contentBlock({
      id: "clause",
      text: "Clause",
      containerPath: [{ kind: "section", id: "schedule-a" }],
    });
    const revised = contentBlock({
      id: "clause",
      text: "Clause",
      containerPath: [{ kind: "section", id: "schedule-b" }],
    });

    const comparison = successfulComparison({ base: [base], revised: [revised] });

    expect(eventTypes(comparison)).toEqual(["movedTo", "movedFrom"]);
    expect(comparison.events[0]).toMatchObject({
      type: "movedTo",
      move: {
        id: 1,
        relation: {
          relationType: "whole",
          base: { block: base },
          revised: { block: revised },
          blockChanges: [
            {
              field: "containerPath",
              base: base.containerPath,
              revised: revised.containerPath,
            },
          ],
        },
      },
    });
    expect(comparison.events[1]).toMatchObject({
      type: "movedFrom",
      move: { id: 1, relation: { base: { block: base } } },
    });
  });
});

describe("identity semantics and input boundaries", () => {
  test("one-sided authoritative identity cannot be weakened by a matching hint", () => {
    const authoritative = contentBlock({ id: "shared", text: "Same" });
    const hinted = contentBlock({
      id: "shared",
      text: "Same",
      identitySemantics: "persistent-hint",
    });

    const comparison = successfulComparison({ base: [authoritative], revised: [hinted] });

    expect(eventTypes(comparison)).toEqual(["inserted", "deleted"]);
    expect(baseProjection(comparison)).toEqual([authoritative]);
    expect(revisedProjection(comparison)).toEqual([hinted]);
  });

  test("stable IDs survive a shift while positional IDs follow their content", () => {
    const stableBase = [
      contentBlock({ id: "a", text: "Repeated" }),
      contentBlock({ id: "b", text: "Repeated" }),
    ];
    const stableInserted = contentBlock({ id: "new", text: "Repeated" });
    const stableRevised = [
      stableInserted,
      contentBlock({ id: "a", text: "Repeated" }),
      contentBlock({ id: "b", text: "Repeated" }),
    ];
    const stable = successfulComparison({ base: stableBase, revised: stableRevised });

    expect(
      stable.events.map((event) => ({
        type: event.type,
        baseId: eventBaseBlocks(event)[0]?.identity.id,
        revisedId: eventRevisedBlocks(event)[0]?.identity.id,
      })),
    ).toEqual([
      { type: "inserted", baseId: undefined, revisedId: "new" },
      { type: "unchanged", baseId: "a", revisedId: "a" },
      { type: "unchanged", baseId: "b", revisedId: "b" },
    ]);

    const positionalBase = [
      contentBlock({ id: "0", text: "Alpha", identitySemantics: "positional" }),
      contentBlock({ id: "1", text: "Beta", identitySemantics: "positional" }),
    ];
    const positionalRevised = [
      contentBlock({ id: "0", text: "Inserted", identitySemantics: "positional" }),
      contentBlock({ id: "1", text: "Alpha", identitySemantics: "positional" }),
      contentBlock({ id: "2", text: "Beta", identitySemantics: "positional" }),
    ];
    const positional = successfulComparison({
      base: positionalBase,
      revised: positionalRevised,
    });

    expect(
      positional.events.map((event) => ({
        type: event.type,
        baseId: eventBaseBlocks(event)[0]?.identity.id,
        revisedId: eventRevisedBlocks(event)[0]?.identity.id,
      })),
    ).toEqual([
      { type: "inserted", baseId: undefined, revisedId: "0" },
      { type: "unchanged", baseId: "0", revisedId: "1" },
      { type: "unchanged", baseId: "1", revisedId: "2" },
    ]);
  });

  test("duplicate IDs return a typed error with the exact side and block index", () => {
    const result = compareContent({
      base: {
        blocks: [
          contentBlock({ id: "duplicate", text: "First" }),
          contentBlock({ id: "duplicate", text: "Second" }),
        ],
      },
      revised: { blocks: [] },
    });

    expect(result.isErr()).toBe(true);
    if (!result.isErr()) {
      return;
    }
    expect(result.error).toBeInstanceOf(InvalidFolioContentComparisonError);
    expect(result.error).toMatchObject({
      input: "base",
      blockIndex: 1,
      field: "blocks[1].identity.id",
    });
  });

  test("one structural identity cannot name multiple coordinates or vice versa", () => {
    const sharedCoordinate = [
      tableBlock({ id: "first", text: "A", rowIndex: 0, cellIndex: 0 }),
      tableBlock({
        id: "second",
        text: "B",
        rowIndex: 0,
        cellIndex: 0,
        paragraphIndex: 1,
        cellIdentity: { type: "authoritative", id: "different-cell" },
      }),
    ];
    const sharedIdentity = [
      tableBlock({
        id: "first",
        text: "A",
        rowIndex: 0,
        cellIndex: 0,
        cellIdentity: { type: "authoritative", id: "shared-cell" },
      }),
      tableBlock({
        id: "second",
        text: "B",
        rowIndex: 0,
        cellIndex: 1,
        cellIdentity: { type: "authoritative", id: "shared-cell" },
      }),
    ];

    for (const blocks of [sharedCoordinate, sharedIdentity]) {
      const result = compareContent({ base: { blocks }, revised: { blocks: [] } });
      expect(result.isErr()).toBe(true);
      if (!result.isErr()) continue;
      expect(result.error).toMatchObject({
        input: "base",
        blockIndex: 1,
        field: "blocks[1].table.cellIdentity",
      });
    }
  });

  test("malformed optional projections return field-specific errors", () => {
    const cases: {
      block: TestBlock;
      field: string;
    }[] = [
      {
        block: contentBlock({
          id: "runs",
          text: "Whole text",
          runs: [{ text: "Partial" }],
        }),
        field: "blocks[0].runs",
      },
      {
        block: contentBlock({
          id: "container",
          text: "Text",
          containerPath: [{ kind: "", id: "section" }],
        }),
        field: "blocks[0].containerPath[0].kind",
      },
      {
        block: tableBlock({
          id: "table",
          text: "Text",
          rowIndex: 0,
          cellIndex: 0,
          columnSpan: 0,
        }),
        field: "blocks[0].table.columnSpan",
      },
    ];

    for (const { block, field } of cases) {
      const result = compareContent({
        base: { blocks: [block] },
        revised: { blocks: [] },
      });
      expect(result.isErr()).toBe(true);
      if (!result.isErr()) {
        continue;
      }
      expect(result.error).toBeInstanceOf(InvalidFolioContentComparisonError);
      expect(result.error).toMatchObject({ input: "base", blockIndex: 0, field });
    }
  });

  test("unrecognized runtime metadata cannot masquerade as declared formatting", () => {
    const base = contentBlock({
      id: "formatting",
      text: "Text",
      runs: [{ text: "Text" }],
    });
    Reflect.set(base, "styleId", 42);
    Reflect.set(base, "displayLabel", false);
    Reflect.set(base, "directSpacing", "120");
    const run = base.runs.at(0);
    if (!run) {
      throw new Error("The malformed-formatting fixture must contain one preview run.");
    }
    Reflect.set(run, "directFormatting", "bold");
    const revised = contentBlock({ id: "formatting", text: "Text", runs: [{ text: "Text" }] });

    const comparison = successfulComparison({ base: [base], revised: [revised] });

    expect(eventTypes(comparison)).toEqual(["unchanged"]);
  });

  test("an unsupported runtime granularity returns an option error", () => {
    const baseBlocks: TestBlock[] = [];
    const revisedBlocks: TestBlock[] = [];
    const options = {
      base: { blocks: baseBlocks },
      revised: { blocks: revisedBlocks },
    };
    Reflect.set(options, "granularity", "byte");

    const result = compareContent(options);

    expect(result.isErr()).toBe(true);
    if (!result.isErr()) {
      return;
    }
    expect(result.error).toBeInstanceOf(InvalidFolioContentComparisonError);
    expect(result.error).toMatchObject({ input: "options", field: "granularity" });
  });

  test("a malformed options value returns a typed error instead of throwing", () => {
    const result = Reflect.apply(compareContent, undefined, [null]);

    expect(result.isErr()).toBe(true);
    if (!result.isErr()) {
      return;
    }
    expect(result.error).toBeInstanceOf(InvalidFolioContentComparisonError);
    expect(result.error).toMatchObject({ input: "options", field: "options" });
  });

  test("public block and change caps return typed limit errors", () => {
    const oneBlock = contentBlock({ id: "same", text: "Text" });
    const tooManyBaseBlocks = Array.from(
      { length: FOLIO_CONTENT_COMPARISON_LIMITS.blocksPerSnapshot + 1 },
      () => oneBlock,
    );
    const blockLimit = compareContent({
      base: { blocks: tooManyBaseBlocks },
      revised: { blocks: [] },
    });

    expect(blockLimit.isErr()).toBe(true);
    if (!blockLimit.isErr()) {
      return;
    }
    expect(blockLimit.error).toBeInstanceOf(FolioContentComparisonLimitError);
    expect(blockLimit.error).toMatchObject({
      input: "base",
      limit: "blocksPerSnapshot",
      maximum: FOLIO_CONTENT_COMPARISON_LIMITS.blocksPerSnapshot,
      actual: FOLIO_CONTENT_COMPARISON_LIMITS.blocksPerSnapshot + 1,
    });

    const tooManyChanges = Array.from(
      { length: FOLIO_CONTENT_COMPARISON_LIMITS.changes + 1 },
      (_unused, index) => contentBlock({ id: `inserted-${String(index)}`, text: "Text" }),
    );
    const changeLimit = compareContent({
      base: { blocks: [] },
      revised: { blocks: tooManyChanges },
    });

    expect(changeLimit.isErr()).toBe(true);
    if (!changeLimit.isErr()) {
      return;
    }
    expect(changeLimit.error).toBeInstanceOf(FolioContentComparisonLimitError);
    expect(changeLimit.error).toMatchObject({
      input: "result",
      limit: "changes",
      maximum: FOLIO_CONTENT_COMPARISON_LIMITS.changes,
      actual: FOLIO_CONTENT_COMPARISON_LIMITS.changes + 1,
    });
  });

  test("comparison is deterministic and does not mutate frozen inputs", () => {
    const base: FolioContentSnapshot = {
      blocks: [
        contentBlock({
          id: "a",
          text: "Alpha beta",
          runs: [{ text: "Alpha ", bold: true }, { text: "beta" }],
          containerPath: [{ kind: "section", id: "main" }],
        }),
      ],
    };
    const revised: FolioContentSnapshot = {
      blocks: [
        contentBlock({
          id: "a",
          text: "Alpha gamma",
          runs: [{ text: "Alpha " }, { text: "gamma", italic: true }],
          containerPath: [{ kind: "section", id: "main" }],
        }),
      ],
    };
    const before = JSON.stringify({ base, revised });
    for (const snapshot of [base, revised]) {
      for (const block of snapshot.blocks) {
        block.runs?.forEach((run) => {
          if (run.effectiveFormatting) Object.freeze(run.effectiveFormatting);
          if (run.authoredFormatting) Object.freeze(run.authoredFormatting);
          Object.freeze(run);
        });
        block.containerPath?.forEach(Object.freeze);
        if (block.runs) Object.freeze(block.runs);
        if (block.containerPath) Object.freeze(block.containerPath);
        Object.freeze(block);
      }
      Object.freeze(snapshot.blocks);
      Object.freeze(snapshot);
    }

    const first = compareContent({ base, revised });
    const second = compareContent({ base, revised });

    expect(first.isErr()).toBe(false);
    expect(second.isErr()).toBe(false);
    if (first.isErr() || second.isErr()) {
      return;
    }
    expect(first.value).toEqual(second.value);
    expect(JSON.stringify({ base, revised })).toBe(before);
  });

  test("every public result node and leaf is recursively runtime-immutable", () => {
    const table = tableBlock({ id: "clause", text: "Before", rowIndex: 0, cellIndex: 0 }).table;
    if (!table) throw new Error("the immutable-result fixture requires a table location");
    const base = contentBlock({
      id: "clause",
      text: "Before",
      kind: "heading",
      headingLevel: 1,
      styleId: "Base",
      directSpacing: { spaceAfter: 120 },
      runs: [{ text: "Before", bold: true, directFormatting: { color: "112233" } }],
      structuralBoundaries: [{ type: "pageBreak", offset: 0, clear: "all" }],
      table,
      containerPath: [{ kind: "section", id: "schedule" }],
    });
    const revised = contentBlock({
      id: "clause",
      text: "After",
      runs: [{ text: "After", italic: true }],
      table,
      containerPath: [{ kind: "section", id: "schedule" }],
    });
    const paired = successfulComparison({ base: [base], revised: [revised] });
    const structural = successfulComparison({
      base: [],
      revised: [tableBlock({ id: "inserted", text: "Inserted", rowIndex: 0, cellIndex: 0 })],
    });

    expectRecursivelyFrozen(paired);
    expectRecursivelyFrozen(structural);
    const pairedBlock = paired.events.at(0);
    if (pairedBlock?.type !== "modified") throw new Error("expected a modified event");
    expect(Reflect.set(pairedBlock.relation.base.block, "text", "mutated")).toBe(false);
  });
});

const generatedText = fc
  .array(fc.constantFrom("alpha", " ", "beta", "\n", "§", "😀", "č", "م", "e\u0301"), {
    maxLength: 8,
  })
  .map((parts) => parts.join(""));

const generatedBlocks = fc.array(
  fc.record({
    kind: fc.constantFrom<TestBlockKind>("paragraph", "heading"),
    text: generatedText,
  }),
  { maxLength: 14 },
);

const generatedMutations = fc.array(
  fc.record({
    text: generatedText,
    kind: fc.constantFrom<TestBlockKind>("paragraph", "heading"),
    bold: fc.boolean(),
    color: fc.constantFrom("112233", "445566", "theme-accent"),
    styleId: fc.constantFrom("Body", "Quote", "Heading1"),
    alignment: fc.constantFrom("left", "center", "start"),
    boundary: fc.constantFrom<"end" | "none" | "start">("none", "start", "end"),
    container: fc.option(fc.constantFrom("body", "schedule-a", "schedule-b"), {
      nil: undefined,
    }),
    mutation: fc.constantFrom<
      | "authored-paragraph"
      | "block"
      | "container"
      | "delete"
      | "effective-paragraph"
      | "inline"
      | "structure"
      | "text"
      | "unchanged"
    >(
      "unchanged",
      "text",
      "block",
      "authored-paragraph",
      "effective-paragraph",
      "inline",
      "structure",
      "container",
      "delete",
    ),
    insertAfter: fc.boolean(),
    moveRank: fc.integer({ min: 0, max: 20 }),
  }),
  { maxLength: 12 },
);

const generatedTableRows = fc.array(fc.array(generatedText, { minLength: 1, maxLength: 3 }), {
  maxLength: 3,
});

const generatedSplitPart = fc.constantFrom(
  "alpha",
  "defined term",
  "§ 12",
  "😀 clause",
  "článek pět",
  "بند قانوني",
);

const boundaryAt = (
  placement: "end" | "none" | "start",
  text: string,
): TestBlock["structuralBoundaries"] =>
  placement === "none"
    ? []
    : [{ type: "pageBreak", offset: placement === "start" ? 0 : text.length, clear: "none" }];

const stablePermutation = fc.shuffledSubarray(
  ["stable-a", "stable-b", "stable-c", "stable-d", "stable-e", "stable-f"],
  { minLength: 6, maxLength: 6 },
);

describe("comparison projection invariants", () => {
  test("every generated stream reconstructs both ordered inputs exactly", () => {
    fc.assert(
      fc.property(generatedBlocks, generatedBlocks, (baseValues, revisedValues) => {
        const base = baseValues.map(({ kind, text }, index) =>
          contentBlock({
            id: `base-${String(index)}`,
            kind,
            text,
            identitySemantics: "positional",
          }),
        );
        const revised = revisedValues.map(({ kind, text }, index) =>
          contentBlock({
            id: `revised-${String(index)}`,
            kind,
            text,
            identitySemantics: "positional",
          }),
        );
        const comparison = successfulComparison({ base, revised });

        expectComparisonReconstructs(comparison, base, revised);
      }),
      propertyConfig({ numRuns: 200 }),
    );
  });

  test("generated semantic mutations reconstruct text, formatting, structure, and identity", () => {
    fc.assert(
      fc.property(generatedMutations, (mutations) => {
        const base = mutations.map(
          ({ text, kind, bold, color, styleId, alignment, boundary, container }, index) =>
            contentBlock({
              id: `stable-${String(index)}`,
              identitySemantics: "authoritative",
              text,
              kind,
              headingLevel: kind === "heading" ? 1 : undefined,
              styleId,
              effectiveParagraphFormatting: { alignment },
              runs:
                text.length === 0
                  ? []
                  : [{ text, bold, directFormatting: { color } }],
              structuralBoundaries: boundaryAt(boundary, text),
              containerPath:
                container === undefined ? [] : [{ kind: "section", id: container }],
            }),
        );
        const revised = mutations
          .map((entry, index) => {
            const revisedText = entry.mutation === "text" ? `${entry.text} Δ` : entry.text;
            const revisedKind =
              entry.mutation === "block"
                ? entry.kind === "heading"
                  ? "paragraph"
                  : "heading"
                : entry.kind;
            const survivor =
              entry.mutation === "delete"
                ? []
                : [
                    contentBlock({
                      id: `stable-${String(index)}`,
                      identitySemantics: "authoritative",
                      text: revisedText,
                      kind: revisedKind,
                      headingLevel: revisedKind === "heading" ? 1 : undefined,
                      styleId:
                        entry.mutation === "authored-paragraph"
                          ? `${entry.styleId}-revised`
                          : entry.styleId,
                      effectiveParagraphFormatting: {
                        alignment:
                          entry.mutation === "effective-paragraph"
                            ? `${entry.alignment}-revised`
                            : entry.alignment,
                      },
                      runs:
                        revisedText.length === 0
                          ? []
                          : [
                              {
                                text: revisedText,
                                bold: entry.mutation === "inline" ? !entry.bold : entry.bold,
                                directFormatting: {
                                  color:
                                    entry.mutation === "inline"
                                      ? `${entry.color}-revised`
                                      : entry.color,
                                },
                              },
                            ],
                      structuralBoundaries: boundaryAt(
                        entry.mutation === "structure"
                          ? entry.boundary === "none"
                            ? "start"
                            : "none"
                          : entry.boundary,
                        revisedText,
                      ),
                      containerPath:
                        entry.mutation === "container"
                          ? [{ kind: "section", id: `${entry.container ?? "body"}-revised` }]
                          : entry.container === undefined
                            ? []
                            : [{ kind: "section", id: entry.container }],
                    }),
                  ];
            const inserted = entry.insertAfter
              ? [
                  contentBlock({
                    id: `inserted-${String(index)}`,
                    identitySemantics: "authoritative",
                    text: `Inserted ${entry.text}`,
                    runs: [{ text: `Inserted ${entry.text}`, italic: true }],
                  }),
                ]
              : [];
            return { rank: entry.moveRank, index, blocks: [...survivor, ...inserted] };
          })
          .toSorted((left, right) => left.rank - right.rank || left.index - right.index)
          .flatMap(({ blocks }) => blocks);

        const comparison = successfulComparison({ base, revised });

        expectComparisonReconstructs(comparison, base, revised);
      }),
      propertyConfig({ numRuns: 160 }),
    );
  });

  test("generated table stories reconstruct row-major structural members", () => {
    fc.assert(
      fc.property(generatedTableRows, generatedTableRows, (baseRows, revisedRows) => {
        const toTableBlocks = (rows: readonly (readonly string[])[], side: string): TestBlock[] =>
          rows.flatMap((cells, rowIndex) =>
            cells.map((text, cellIndex) =>
              tableBlock({
                id: `${side}-${String(rowIndex)}-${String(cellIndex)}`,
                text,
                rowIndex,
                cellIndex,
              }),
            ),
          );
        const base = toTableBlocks(baseRows, "base");
        const revised = toTableBlocks(revisedRows, "revised");

        const comparison = successfulComparison({ base, revised });

        expectComparisonReconstructs(comparison, base, revised);
      }),
      propertyConfig({ numRuns: 100 }),
    );
  });

  test("generated splits and merges reconstruct either stable survivor direction", () => {
    fc.assert(
      fc.property(
        generatedSplitPart,
        generatedSplitPart,
        fc.constantFrom(" ", "\n", "\t", "\r\n"),
        fc.boolean(),
        (leftText, rightText, separator, rightSurvives) => {
          const joinedId = rightSurvives ? "right" : "left";
          const joinedText = `${leftText}${separator}${rightText}`;
          const joinedBase = contentBlock({ id: joinedId, text: joinedText });
          const splitRevised = [
            contentBlock({ id: "left", text: leftText }),
            contentBlock({ id: "right", text: rightText }),
          ];
          const split = successfulComparison({ base: [joinedBase], revised: splitRevised });
          expect(eventTypes(split)).toEqual(["split"]);
          expectComparisonReconstructs(split, [joinedBase], splitRevised);

          const mergeBase = [
            contentBlock({ id: "left", text: leftText }),
            contentBlock({ id: "right", text: rightText }),
          ];
          const joinedRevised = contentBlock({ id: joinedId, text: joinedText });
          const merge = successfulComparison({ base: mergeBase, revised: [joinedRevised] });
          expect(eventTypes(merge)).toEqual(["merge"]);
          expectComparisonReconstructs(merge, mergeBase, [joinedRevised]);
        },
      ),
      propertyConfig({ numRuns: 100 }),
    );
  });

  test("permutations of short stable blocks preserve projections and move identity", () => {
    fc.assert(
      fc.property(stablePermutation, (revisedIds) => {
        const baseIds = ["stable-a", "stable-b", "stable-c", "stable-d", "stable-e", "stable-f"];
        const base = baseIds.map((id) => contentBlock({ id, text: id.at(-1) ?? id }));
        const revised = revisedIds.map((id) => contentBlock({ id, text: id.at(-1) ?? id }));

        const comparison = successfulComparison({ base, revised });

        expect(baseProjection(comparison)).toEqual(base);
        expect(revisedProjection(comparison)).toEqual(revised);
        expect(
          comparison.events.every(
            ({ type }) => type === "unchanged" || type === "movedFrom" || type === "movedTo",
          ),
        ).toBe(true);
        const sourceByMoveId = new Map<number, string>();
        const targetByMoveId = new Map<number, string>();
        for (const event of comparison.events) {
          if (event.type === "movedFrom") {
            sourceByMoveId.set(event.move.id, event.move.relation.base.block.identity.id);
          }
          if (event.type === "movedTo") {
            expect(event.move.relation.base.block.identity.id).toBe(event.move.relation.revised.block.identity.id);
            targetByMoveId.set(event.move.id, event.move.relation.base.block.identity.id);
          }
        }
        expect(targetByMoveId).toEqual(sourceByMoveId);
      }),
      propertyConfig({ numRuns: 100 }),
    );
  });
});
