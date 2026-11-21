import { expect, spyOn, test } from "bun:test";

import { assertExactModel } from "../../../../../test/exactModel";
import type { Document, Paragraph } from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import * as blocks from "../blocks";
import { tryTrackedInsertionGroup } from "../documentOps";
import * as edits from "../edits";
import * as ids from "../ids";
import * as primitives from "../paragraphPrimitive";
import {
  getParagraphPropertySnapshot,
  registerParagraphPropertySnapshot,
} from "../paragraphPropertySnapshot";
import { DOCUMENT_OP_TYPES, type InsertBlocksOp, OP_STORIES } from "../types";

const paragraph = (paraId: string, text: string): Paragraph => ({
  type: "paragraph",
  paraId,
  content:
    text === "" ? [] : [{ type: "run", content: [{ type: "text", text }] }],
});

const sourceId = "10000001";
const makeSource = ({
  following,
  topology,
  tracked,
}: {
  following: boolean;
  topology: "trailing" | "stale";
  tracked?: boolean;
}) => {
  const source = paragraph(sourceId, "original");
  const anchor: Paragraph = tracked
    ? {
        ...source,
        content: [
          {
            type: "insertion",
            info: { id: 17, author: "Earlier" },
            content: source.content,
          },
        ],
      }
    : source;
  const follower = paragraph("10000002", "follower");
  registerParagraphPropertySnapshot({
    paragraph: anchor,
    snapshot: { alignment: "center" },
  });
  const content = following ? [anchor, follower] : [anchor];
  const properties = {};
  const document = {
    package: {
      document: {
        content,
        finalSectionProperties: properties,
        sections: [
          {
            properties,
            content:
              topology === "trailing"
                ? [...content]
                : [paragraph(sourceId, "wrong")],
          },
        ],
      },
    },
  } satisfies Document;
  return { anchor, document };
};

type MemberOptions = {
  side: "before" | "after";
  index: number;
  previous: readonly number[] | undefined;
};
const compile = ({ side, index, previous }: MemberOptions): InsertBlocksOp => {
  const id = Math.max(99, ...(previous ?? [])) + 1;
  const blockId = (0x20000000 + index).toString(16);
  const draft = paragraph(blockId, index % 3 === 0 ? "" : `insert ${index}`);
  return {
    type: DOCUMENT_OP_TYPES.INSERT_BLOCKS,
    story: OP_STORIES.MAIN,
    at: { type: side, blockId: sourceId },
    blocks: [
      index === 0 ? { ...draft, formatting: { alignment: "right" } } : draft,
    ],
    revision: { id, author: "Reviewer", date: "2026-01-11T00:00:00Z" },
    newIds: { revision: [id + 1, id + 2, id + 3, id + 4] },
  };
};

const snapshots = (document: Document) =>
  blocks
    .storyParagraphs(document.package.document)
    .map(({ paragraph: block }) => [
      block.paraId,
      getParagraphPropertySnapshot(block),
    ]);

test.each([
  [4, "before", "tail"],
  [16, "after", "tail"],
  [64, "after", "follower"],
] as const)(
  "%i tracked insertions (%s, %s) materialize once with exact member effects",
  (count, side, following) => {
    const { anchor, document } = makeSource({
      following: following === "follower",
      topology: "trailing",
      tracked: side === "after",
    });
    const baseModel = structuredClone(document);
    const splices = spyOn(blocks, "replaceParagraphs");
    const combine = spyOn(edits, "combineEditEffects");
    try {
      let nextIndex = 0;
      const grouped = tryTrackedInsertionGroup({
        document,
        next: (previous) => {
          if (nextIndex === count) return { type: "done" };
          const index = nextIndex;
          nextIndex += 1;
          return { type: "member", op: compile({ index, side, previous }) };
        },
      });
      if (grouped.type !== "materialized")
        throw new Error(JSON.stringify(grouped));
      expect(splices).toHaveBeenCalledTimes(1);
      expect(combine).toHaveBeenCalledTimes(0);
      splices.mockClear();
      let ordinary = document;
      let previous: readonly number[] | undefined;
      const prior: Document[] = [];
      for (let index = 0; index < count; index++) {
        const op = compile({ index, side, previous });
        const { document: output, ...expected } = applyDocumentOp(
          ordinary,
          op,
        ).unwrap();
        const member = grouped.members.at(index);
        expect(member).toBeDefined();
        assertExactModel(member, { op, ...expected });
        prior.push(ordinary);
        ordinary = output;
        previous = expected.revisions;
      }
      expect(splices.mock.calls.length).toBeGreaterThanOrEqual(count);
      expect(grouped.members).toHaveLength(count);
      assertExactModel(grouped.document, ordinary);
      assertExactModel(snapshots(grouped.document), snapshots(ordinary));
      expect(getParagraphPropertySnapshot(anchor)).toEqual({
        alignment: "center",
      });
      assertExactModel(document, baseModel);

      let undone = grouped.document;
      for (const [index, member] of grouped.members
        .entries()
        .toArray()
        .toReversed()) {
        undone = applyDocumentOps(undone, member.inverse).unwrap().document;
        assertExactModel(undone, prior.at(index));
      }
      assertExactModel(undone, document);
    } finally {
      combine.mockRestore();
      splices.mockRestore();
    }
  },
);

test.each(["paragraph id", "revision id"] as const)(
  "group abandons the entire draft on a later duplicate %s",
  (collision) => {
    const { document } = makeSource({ following: false, topology: "trailing" });
    const base = structuredClone(document);
    const op = compile({ side: "after", index: 0, previous: undefined });
    const duplicate =
      collision === "paragraph id"
        ? compile({ side: "after", index: 0, previous: [200] })
        : compile({ side: "after", index: 1, previous: undefined });
    const ordinary = applyDocumentOp(document, op).unwrap();
    const expected = applyDocumentOp(ordinary.document, duplicate);
    if (expected.isOk()) throw new Error("Expected a native collision");
    const splices = spyOn(blocks, "replaceParagraphs");
    try {
      const group = tryTrackedInsertionGroup({
        document,
        next: (previous) =>
          previous === undefined
            ? { type: "member", op }
            : { type: "member", op: duplicate },
      });
      if (group.type !== "fallback" || group.reason !== "refusal")
        throw new Error("Expected a later native refusal");
      expect(group.index).toBe(1);
      assertExactModel({ ...group.refusal }, { ...expected.error });
      expect(splices).toHaveBeenCalledTimes(0);
      assertExactModel(document, base);
    } finally {
      splices.mockRestore();
    }
  },
);

test("an insertion before a retained paragraph refuses its occupied id with native precedence", () => {
  const { document } = makeSource({
    following: false,
    topology: "trailing",
    tracked: true,
  });
  const drafted = compile({ side: "before", index: 1, previous: undefined });
  for (const block of [
    paragraph(sourceId, "new"),
    { ...paragraph(sourceId, "new"), content: [{ type: "run", content: [] }] },
  ] satisfies Paragraph[]) {
    const op = { ...drafted, blocks: [block] };
    const ordinary = applyDocumentOp(document, op);
    if (ordinary.isOk())
      throw new Error("Expected the insertion to be refused");
    const grouped = tryTrackedInsertionGroup({
      document,
      next: () => ({ type: "member", op }),
    });
    if (grouped.type !== "fallback" || grouped.reason !== "refusal")
      throw new Error("Expected the native refusal");
    expect(grouped.index).toBe(0);
    assertExactModel({ ...grouped.refusal }, { ...ordinary.error });
  }
});

test("group falls back for a stale physical section and a compiler abort", () => {
  const op = compile({ side: "after", index: 0, previous: undefined });
  const { document: stale } = makeSource({
    following: false,
    topology: "stale",
  });
  const { document: sound } = makeSource({
    following: false,
    topology: "trailing",
  });
  const next = (previous: readonly number[] | undefined) =>
    previous === undefined
      ? ({ type: "member", op } as const)
      : ({ type: "fallback" } as const);
  const splices = spyOn(blocks, "replaceParagraphs");
  try {
    const rejected = tryTrackedInsertionGroup({ document: stale, next });
    expect(rejected).toEqual({ type: "fallback", reason: "sectionView" });
    expect(tryTrackedInsertionGroup({ document: sound, next })).toEqual({
      type: "fallback",
      reason: "compiler",
    });
    expect(splices).toHaveBeenCalledTimes(0);
  } finally {
    splices.mockRestore();
  }
});

test.each([
  [16_383, "materialized"],
  [16_384, "fallback"],
] as const)(
  "old native splice bounds a tracked group of %i members",
  (count, outcome) => {
    const { document } = makeSource({ following: false, topology: "trailing" });
    let index = 0;
    const splices = spyOn(blocks, "replaceParagraphs");
    try {
      const result = tryTrackedInsertionGroup({
        document,
        next: (previous) =>
          index === count
            ? { type: "done" }
            : {
                type: "member",
                op: compile({ side: "after", index: index++, previous }),
              },
      });
      expect(result.type).toBe(outcome);
      if (result.type === "fallback") {
        expect(result.reason).toBe("wideSpan");
        expect(splices).toHaveBeenCalledTimes(0);
      } else {
        expect(result.document.package.document.content).toHaveLength(
          count + 1,
        );
        expect(splices).toHaveBeenCalledTimes(1);
      }
    } finally {
      splices.mockRestore();
    }
  },
);

test.each([8, 32])(
  "an unchanged neighbor contributes no identity census for %i insertions",
  (count) => {
    const { document, anchor } = makeSource({
      following: false,
      topology: "trailing",
      tracked: true,
    });
    const records = spyOn(ids, "identityKeysIn");
    const paragraphs = spyOn(ids, "paragraphIdsIn");
    try {
      let index = 0;
      const group = tryTrackedInsertionGroup({
        document,
        next: (previous) =>
          index === count
            ? { type: "done" }
            : {
                type: "member",
                op: compile({ side: "before", index: index++, previous }),
              },
      });
      expect(group.type).toBe("materialized");
      const includesAnchor = ([input]: readonly [unknown]) =>
        Array.isArray(input) && input.includes(anchor);
      expect(records.mock.calls.filter(includesAnchor)).toHaveLength(0);
      expect(paragraphs.mock.calls.filter(includesAnchor)).toHaveLength(0);
    } finally {
      records.mockRestore();
      paragraphs.mockRestore();
    }
  },
);

test.each(["tail", "follower"] as const)(
  "draft and stamped insertions use full replacement checks for a changed %s neighbor",
  (placement) => {
    const { document } = makeSource({
      following: placement === "follower",
      topology: "trailing",
    });
    const validator = spyOn(primitives, "validateBlockReplacement");
    try {
      let index = 0;
      const group = tryTrackedInsertionGroup({
        document,
        next: (previous) =>
          index === 8
            ? { type: "done" }
            : {
                type: "member",
                op: compile({ side: "after", index: index++, previous }),
              },
      });
      expect(group.type).toBe("materialized");
      expect(validator.mock.calls).toHaveLength(16);
      expect(
        validator.mock.calls.filter(([{ before }]) => before.length > 0),
      ).toHaveLength(placement === "tail" ? 2 : 0);
    } finally {
      validator.mockRestore();
    }
  },
);
