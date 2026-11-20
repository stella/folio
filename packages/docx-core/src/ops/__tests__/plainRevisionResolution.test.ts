import { expect, spyOn, test } from "bun:test";

import { assertExactModel } from "../../../../../test/exactModel";
import type { Document, Paragraph, Run } from "../../model/document";
import {
  type AppliedDocumentOp,
  applyDocumentOp,
  applyDocumentOps,
} from "../apply";
import * as blocks from "../blocks";
import {
  getParagraphPropertySnapshot,
  registerParagraphPropertySnapshot,
} from "../paragraphPropertySnapshot";
import { DOCUMENT_OP_REFUSAL_REASONS } from "../refusal";
import * as materialization from "../resolutionMaterialization";
import { planRevisionResolution } from "../resolve";
import {
  DOCUMENT_OP_TYPES,
  OP_STORIES,
  type ResolveRevisionOp,
} from "../types";

const run = (text: string, formatting?: Run["formatting"]): Run => ({
  type: "run",
  content: [{ type: "text", text }],
  ...(formatting === undefined ? {} : { formatting }),
});

const original: Paragraph = {
  type: "paragraph",
  paraId: "10000001",
  content: [run("prefix"), run("bold", { bold: true }), run("suffix")],
};

const bodyOf = (content: Paragraph[]): Document => ({
  package: { document: { content } },
});

const markInfo = (id: number) => ({
  id,
  author: "Reviewer",
  date: "2026-01-11T10:15:30Z",
});

const apply = (
  document: Document,
  input: Parameters<typeof applyDocumentOp>[1],
) => {
  const result = applyDocumentOp(document, input);
  if (result.isErr()) throw result.error;
  return result.value;
};

type TrackedTailOptions = {
  topology?: "sectionless" | "trailing section";
  source?: "empty" | "mixed";
};

const trackedTail = (
  count: number,
  { topology = "sectionless", source = "mixed" }: TrackedTailOptions = {},
) => {
  const anchor = source === "empty" ? { ...original, content: [] } : original;
  let document = bodyOf([anchor]);
  const ids: number[] = [];
  for (let index = 0; index < count; index += 1) {
    const inserted = {
      type: "paragraph",
      paraId: (0x20000000 + index).toString(16),
      content: [run(`inserted ${index} `), run("contrast", { italic: true })],
    } satisfies Paragraph;
    const id = 100 + index * 8;
    const applied = apply(document, {
      type: DOCUMENT_OP_TYPES.INSERT_BLOCKS,
      story: OP_STORIES.MAIN,
      at: { type: "after", blockId: "10000001" },
      blocks: [inserted],
      revision: markInfo(id),
      newIds: { revision: [id + 1, id + 2, id + 3, id + 4] },
    });
    expect(applied.revisions.length).toBeGreaterThan(0);
    document = applied.document;
    ids.push(...applied.revisions);
  }
  const marked = document.package.document.content.map((block) =>
    block.type === "paragraph" ? block.pPrMark : undefined,
  );
  expect(marked.at(0)?.survivor).toBe("first");
  expect(
    marked
      .slice(1, -1)
      .every((mark) => mark?.kind === "ins" && mark.survivor === undefined),
  ).toBe(true);
  expect(marked.at(-1)).toBeUndefined();
  if (topology === "trailing section") {
    // Public projection builds an independent section array and has no
    // section-view certificate from previous native operations.
    const { content } = document.package.document;
    const properties = {};
    document = {
      package: {
        document: {
          content,
          finalSectionProperties: properties,
          sections: [{ properties, content: [...content] }],
        },
      },
    };
  }
  return { document, ids };
};

const resolveOp = (
  decision: ResolveRevisionOp["decision"],
  revisionIds: readonly number[],
): ResolveRevisionOp => ({
  type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
  story: OP_STORIES.MAIN,
  decision,
  revisionIds: [...revisionIds],
});

const snapshotSources = (document: Document) =>
  blocks
    .storyParagraphs(document.package.document)
    .map(({ paragraph }) => [
      paragraph.paraId,
      getParagraphPropertySnapshot(paragraph),
    ]);

const expectExact = (
  actual: AppliedDocumentOp,
  expected: AppliedDocumentOp,
) => {
  assertExactModel(actual.document, expected.document);
  assertExactModel(actual.inverse, expected.inverse);
  assertExactModel(actual.touched, expected.touched);
  assertExactModel(actual.revisions, expected.revisions);
  assertExactModel(
    snapshotSources(actual.document),
    snapshotSources(expected.document),
  );
};

const counted = (document: Document, op: ResolveRevisionOp) => {
  const materialize = spyOn(materialization, "tryMaterializeResolution");
  const splices = spyOn(blocks, "replaceParagraphs");
  try {
    const actual = applyDocumentOp(document, op);
    const successful = materialize.mock.results.filter(
      ({ value }) => value?.type === "materialized",
    ).length;
    const fallbacks = materialize.mock.results.flatMap(({ value }) =>
      value?.type === "fallback" ? [value.reason] : [],
    );
    const actualSplices = splices.mock.calls.length;
    materialize.mockReturnValue({ type: "fallback", reason: "primitiveShape" });
    splices.mockClear();
    const expected = applyDocumentOp(document, op);
    return {
      actual,
      expected,
      successful,
      fallbacks,
      actualSplices,
      generalSplices: splices.mock.calls.length,
    };
  } finally {
    splices.mockRestore();
    materialize.mockRestore();
  }
};

const expectUndo = (edit: AppliedDocumentOp, source: Document) => {
  const undone = applyDocumentOps(edit.document, edit.inverse);
  if (undone.isErr()) throw undone.error;
  assertExactModel(undone.value.document, source);
  assertExactModel(
    snapshotSources(undone.value.document),
    snapshotSources(source),
  );
};

test.each([
  ["accept", 16, "sectionless"],
  ["accept", 48, "trailing section"],
  ["reject", 16, "sectionless"],
  ["reject", 48, "trailing section"],
] as const)(
  "plain tracked tail resolves natively with exact %s effects (%i, %s)",
  (decision, count, topology) => {
    const { document, ids } = trackedTail(count, {
      topology,
      source: decision === "reject" ? "empty" : "mixed",
    });
    const originalModel = structuredClone(document);
    // Resolve with duplicated and reversed request IDs. Effects retain the
    // native plan's document order and only cover actual records.
    const choice = resolveOp(decision, [
      ...ids.toReversed(),
      ...ids.slice(0, 2),
    ]);
    const {
      actual,
      expected,
      successful,
      fallbacks,
      actualSplices,
      generalSplices,
    } = counted(document, choice);
    if (actual.isErr()) throw actual.error;
    if (expected.isErr()) throw expected.error;
    expect(successful).toBe(1);
    expect(fallbacks).toEqual([]);
    expect(actualSplices).toBe(1);
    expect(generalSplices).toBeGreaterThanOrEqual(count);
    expect(generalSplices - actualSplices).toBeGreaterThanOrEqual(count - 1);
    expectExact(actual.value, expected.value);
    expectUndo(actual.value, document);
    assertExactModel(document, originalModel);
  },
);

test("clearing same-format marks carries the authored paragraph snapshot", () => {
  const { document, ids } = trackedTail(16);
  for (const { paragraph } of blocks.storyParagraphs(
    document.package.document,
  )) {
    registerParagraphPropertySnapshot({
      paragraph,
      snapshot: { alignment: "center" },
    });
  }
  const { actual, expected, successful } = counted(
    document,
    resolveOp("accept", ids),
  );
  if (actual.isErr()) throw actual.error;
  if (expected.isErr()) throw expected.error;
  expect(successful).toBe(1);
  expectExact(actual.value, expected.value);
  expect(
    snapshotSources(actual.value.document).map(([, source]) => source),
  ).toEqual(Array.from({ length: 17 }, () => ({ alignment: "center" })));
  expectUndo(actual.value, document);
});

test.each(["accept", "reject"] as const)(
  "an unselected wrapper in a changed paragraph stays on the general %s path",
  (decision) => {
    const make = (position: number): Paragraph => ({
      ...original,
      paraId: (0x10000000 + position).toString(16),
      content: [
        {
          type: "insertion",
          info: markInfo(2 * position),
          content: [run("selected")],
        },
        {
          type: "insertion",
          info: markInfo(2 * position + 1),
          content: [run("pending")],
        },
      ],
    });
    const document = bodyOf(
      Array.from({ length: 16 }, (_, index) => make(index + 1)),
    );
    const ids = Array.from({ length: 16 }, (_, index) => 2 * (index + 1));
    const {
      actual,
      expected,
      successful,
      fallbacks,
      actualSplices,
      generalSplices,
    } = counted(document, resolveOp(decision, ids));
    if (actual.isErr()) throw actual.error;
    if (expected.isErr()) throw expected.error;
    expect(successful).toBe(0);
    expect(fallbacks).toHaveLength(1);
    expect(actualSplices).toBe(generalSplices);
    expect(actualSplices).toBeGreaterThanOrEqual(ids.length);
    expectExact(actual.value, expected.value);
    expectUndo(actual.value, document);
  },
);

test("an invalid requested ID precedes a selected final section record", () => {
  const { document: tracked, ids } = trackedTail(16);
  const finalRevision = 1000;
  const document: Document = {
    ...tracked,
    package: {
      ...tracked.package,
      document: {
        ...tracked.package.document,
        finalSectionProperties: {
          propertyChanges: [
            {
              type: "sectionPropertyChange",
              info: markInfo(finalRevision),
              previousProperties: {},
            },
          ],
        },
      },
    },
  };
  for (const { selected, reason, message } of [
    {
      selected: [finalRevision, ...ids, 1.5],
      reason: DOCUMENT_OP_REFUSAL_REASONS.INVALID_NEW_ID,
      message: "A revision id is an integer.",
    },
    {
      selected: [finalRevision, ...ids],
      reason: DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
      message: `Final section revision ${finalRevision} requires section review.`,
    },
  ]) {
    const { actual, expected, successful, fallbacks } = counted(
      document,
      resolveOp("accept", selected),
    );
    if (actual.isOk() || expected.isOk())
      throw new Error("Expected an atomic refusal");
    expect(successful).toBe(0);
    expect(fallbacks).toEqual([]);
    expect({ ...actual.error }).toEqual({ ...expected.error });
    expect(actual.error.reason).toBe(reason);
    expect(actual.error.message).toBe(message);
  }
});

test("a staged invalid XML character keeps the native refusal and leaves input intact", () => {
  const make = (id: number): Paragraph => ({
    ...original,
    paraId: (0x10000000 + id).toString(16),
    content: [
      {
        type: "insertion",
        info: markInfo(id),
        content: [run(id === 12 ? "bad\u0000text" : "ordinary text")],
      },
    ],
  });
  const ids = Array.from({ length: 16 }, (_, index) => index + 1);
  const document = bodyOf(ids.map(make));
  const originalModel = structuredClone(document);
  const { actual, expected, successful, fallbacks } = counted(
    document,
    resolveOp("accept", ids),
  );
  if (actual.isOk() || expected.isOk())
    throw new Error("Expected an atomic refusal");
  expect(successful).toBe(0);
  expect(fallbacks).toEqual(["primitiveValidation"]);
  expect({ ...actual.error }).toEqual({ ...expected.error });
  expect(actual.error.reason).toBe(DOCUMENT_OP_REFUSAL_REASONS.INVALID_TEXT);
  assertExactModel(document, originalModel);
});

test("declared comments keep the plain-resolution optimization off", () => {
  const { document: tracked, ids } = trackedTail(16);
  const document: Document = {
    ...tracked,
    package: {
      ...tracked.package,
      document: {
        ...tracked.package.document,
        comments: [
          {
            id: 41,
            author: "Reviewer",
            content: [{ ...original, paraId: "30000001" }],
          },
        ],
      },
    },
  };
  const {
    actual,
    expected,
    successful,
    fallbacks,
    actualSplices,
    generalSplices,
  } = counted(document, resolveOp("accept", ids));
  if (actual.isErr()) throw actual.error;
  if (expected.isErr()) throw expected.error;
  expect(successful).toBe(0);
  expect(fallbacks).toEqual(["packageDependencies"]);
  expect(actualSplices).toBe(generalSplices);
  expectExact(actual.value, expected.value);
  expectUndo(actual.value, document);
});

test("a rejected tail with remaining anchor text falls back to the native join path", () => {
  const { document, ids } = trackedTail(16);
  const {
    actual,
    expected,
    successful,
    fallbacks,
    actualSplices,
    generalSplices,
  } = counted(document, resolveOp("reject", ids));
  if (actual.isErr()) throw actual.error;
  if (expected.isErr()) throw expected.error;
  expect(successful).toBe(0);
  expect(fallbacks).toEqual(["joinShape"]);
  expect(actualSplices).toBe(generalSplices);
  expectExact(actual.value, expected.value);
  expectUndo(actual.value, document);
});

test("the eligible reject plan emits the original receivers in the original order", () => {
  const { document, ids } = trackedTail(16, { source: "empty" });
  const options = {
    document,
    op: resolveOp("reject", ids),
    applyOps: applyDocumentOps,
  };
  const materialize = spyOn(materialization, "tryMaterializeResolution");
  try {
    const actual = planRevisionResolution(options).unwrap();
    expect(materialize.mock.results.at(0)?.value?.type).toBe("materialized");
    materialize.mockReturnValue({ type: "fallback", reason: "primitiveShape" });
    const expected = planRevisionResolution(options).unwrap();
    expect(actual.receiverEffects).toHaveLength(16);
    assertExactModel(actual.receiverEffects, expected.receiverEffects);
  } finally {
    materialize.mockRestore();
  }
});

test("a middle accept splices only the selected paragraph", () => {
  const midpoint = 64;
  const selectedId = 999;
  const paragraphs = Array.from({ length: 129 }, (_, index) => ({
    ...original,
    paraId: (0x10000000 + index).toString(16),
    content:
      index === midpoint
        ? [
            {
              type: "insertion" as const,
              info: markInfo(selectedId),
              content: [run("selected")],
            },
          ]
        : original.content,
  }));
  const document = bodyOf(paragraphs);
  const splices = spyOn(blocks, "replaceParagraphs");
  const materialize = spyOn(materialization, "tryMaterializeResolution");
  try {
    const actual = apply(document, resolveOp("accept", [selectedId]));
    expect(materialize.mock.results.at(0)?.value?.type).toBe("materialized");
    expect(splices).toHaveBeenCalledTimes(1);
    const options = splices.mock.calls.at(0)?.at(0);
    expect(options?.at).toEqual({ list: [], index: midpoint });
    expect(options?.count).toBe(1);
    expect(options?.replacement).toHaveLength(1);
    const { content } = actual.document.package.document;
    expect(content).toHaveLength(paragraphs.length);
    for (const [index, paragraph] of paragraphs.entries()) {
      if (index !== midpoint) expect(content[index]).toBe(paragraph);
    }
  } finally {
    splices.mockRestore();
    materialize.mockRestore();
  }
});

test("a reject chain reads the staged paragraphs without cloning the source list", () => {
  const { document, ids } = trackedTail(48, {
    topology: "trailing section",
    source: "empty",
  });
  const slices = spyOn(document.package.document.content, "slice");
  const splices = spyOn(blocks, "replaceParagraphs");
  try {
    const { edit } = planRevisionResolution({
      document,
      op: resolveOp("reject", ids),
      applyOps: applyDocumentOps,
    }).unwrap();
    expect(splices).toHaveBeenCalledTimes(1);
    expect(slices).toHaveBeenCalledTimes(0);
    expect(edit.document.package.document.content).toHaveLength(1);
  } finally {
    splices.mockRestore();
    slices.mockRestore();
  }
});
