import { Result } from "better-result";
import { expect, spyOn, test } from "bun:test";

import { MAX_REVISION_ID, type Paragraph } from "@stll/docx-core/model";
import * as documentOperations from "@stll/docx-core/ops";

import { FolioDocxReviewer } from "../ai-edits/headless";
import type {
  FolioDocumentOperation,
  FolioDocumentOperationMode,
} from "../document-operations";
import { prepareCanonicalDocxInput } from "../docx/canonicalSessionInput";
import { createDocx } from "../docx/rezip";
import { createEmptyDocument } from "../utils/createDocument";
import * as replayIds from "./canonicalInsertionReplayIds";

const paragraph = (id: string, text: string): Paragraph => ({
  type: "paragraph",
  paraId: id,
  content: [{ type: "run", content: [{ type: "text", text }] }],
});

const fixture = async () => {
  const template = createEmptyDocument();
  const content = Array.from({ length: 10 }, (_, index) =>
    paragraph(
      (index + 1).toString(16).padStart(8, "0"),
      index === 0 ? "Anchor" : `Original ${index + 1}`,
    ),
  );
  const input = await createDocx({
    ...template,
    package: {
      ...template.package,
      document: { content },
    },
  });
  return (await prepareCanonicalDocxInput(input)).unwrap().slice().buffer;
};

const open = async (input: ArrayBuffer) => {
  const reviewer = await FolioDocxReviewer.fromBuffer(input.slice(0), {
    session: "canonical",
  });
  const anchor = reviewer
    .snapshot()
    .blocks.find(({ text }) => text === "Anchor");
  const retired = reviewer
    .snapshot()
    .blocks.find(({ text }) => text === "Original 3");
  if (anchor === undefined || retired === undefined)
    throw new Error("Expected fixture identities");
  const deleted = reviewer.applyDocumentOperations(
    {
      version: 1,
      mode: "direct",
      atomic: true,
      operations: [{ type: "deleteBlock", id: "retire", blockId: retired.id }],
    },
    { undefinedReferences: "refuse" },
  );
  expect(deleted.skipped).toEqual([]);
  expect(deleted.applied).toHaveLength(1);
  return { reviewer, anchor };
};

const appendRun = ({
  reviewer,
  anchor,
  count,
  edge,
  mode,
  idSeed,
}: {
  reviewer: FolioDocxReviewer;
  anchor: string;
  count: number;
  edge: "insertBeforeBlock" | "insertAfterBlock";
  mode: FolioDocumentOperationMode;
  idSeed?: number;
}) => {
  const operations = Array.from(
    { length: count },
    (_, index) =>
      ({
        type: edge,
        id: `insert-${index}`,
        blockId: anchor,
        text: `Inserted ${index}`,
      }) satisfies FolioDocumentOperation,
  );
  return reviewer.applyDocumentOperations(
    { version: 1, mode, atomic: idSeed === undefined, operations },
    {
      undefinedReferences: "refuse",
      revisionStamp: { idSeed: idSeed ?? 101, date: "2026-03-01T00:00:00Z" },
    },
  );
};

const withoutSessionHandle = ({
  undoHandle,
  ...rest
}: ReturnType<typeof appendRun>) => ({
  ...rest,
  undoHandle: undoHandle?.type ?? null,
});

test.each([
  ["insertBeforeBlock", "tracked-changes", 12],
  ["insertAfterBlock", "direct", 24],
  ["insertAfterBlock", "tracked-changes", 48],
] as const)(
  "%s (%s) reuses admitted IDs for %i reordered operations",
  async (edge, mode, count) => {
    const input = await fixture();
    const fast = await open(input);
    const slow = await open(input);
    const before = fast.reviewer.snapshot().blocks;
    const format = spyOn(documentOperations, "formatParaId");
    const plan = replayIds.planReplayParagraphIds;
    let beforeReplay = 0;
    const proof = spyOn(replayIds, "planReplayParagraphIds").mockImplementation(
      (rejection, members) => {
        beforeReplay = format.mock.calls.length;
        return plan(rejection, members);
      },
    );
    const compile = spyOn(documentOperations, "compileEditorIntent");
    const compiledOps = () =>
      compile.mock.results.flatMap(({ value }) =>
        value?.isOk() ? value.value.ops : [],
      );
    try {
      const actual = appendRun({
        reviewer: fast.reviewer,
        anchor: fast.anchor.id,
        count,
        edge,
        mode,
      });
      expect(actual.skipped).toEqual([]);
      expect(actual.applied).toHaveLength(count);
      expect(proof).toHaveBeenCalledTimes(1);
      expect(proof.mock.results.at(0)?.value?.type).toBe("admitted");
      const admissionFormats = beforeReplay;
      const actualFormats = format.mock.calls.length - beforeReplay;
      const actualOps = compiledOps();
      for (const op of actualOps) {
        if (op.type !== documentOperations.DOCUMENT_OP_TYPES.INSERT_BLOCKS)
          continue;
        expect(op.blocks.at(0)?.paraId).not.toBe("00000003");
        expect(op.blocks.at(0)?.paraId?.toUpperCase()).not.toBe("0000000A");
      }

      format.mockClear();
      compile.mockClear();
      proof.mockImplementation(() => {
        beforeReplay = format.mock.calls.length;
        return { type: "fallback", reason: "windowShape" };
      });
      const expected = appendRun({
        reviewer: slow.reviewer,
        anchor: slow.anchor.id,
        count,
        edge,
        mode,
      });
      expect(withoutSessionHandle(expected)).toEqual(
        withoutSessionHandle(actual),
      );
      expect(compiledOps()).toStrictEqual(actualOps);
      expect(slow.reviewer.snapshot()).toStrictEqual(fast.reviewer.snapshot());
      const ordinaryFormats = format.mock.calls.length - beforeReplay;
      expect(beforeReplay).toBe(admissionFormats);
      expect(actualFormats).toBeLessThanOrEqual(3 * count);
      expect(ordinaryFormats).toBeGreaterThanOrEqual(count * (count - 1));
      if (actual.undoHandle === null || expected.undoHandle === null)
        throw new Error("Expected a native undo handle");
      expect(
        fast.reviewer.undoDocumentOperations(actual.undoHandle).status,
      ).toBe("undone");
      expect(
        slow.reviewer.undoDocumentOperations(expected.undoHandle).status,
      ).toBe("undone");
      expect(fast.reviewer.snapshot().blocks).toStrictEqual(before);
      expect(slow.reviewer.snapshot().blocks).toStrictEqual(before);
    } finally {
      compile.mockRestore();
      format.mockRestore();
      proof.mockRestore();
    }
  },
);

test("revision exhaustion preserves partially admitted results and stamps", async () => {
  const input = await fixture();
  const fast = await open(input);
  const slow = await open(input);
  const proof = spyOn(replayIds, "planReplayParagraphIds");
  try {
    const request = {
      count: 4,
      edge: "insertBeforeBlock",
      mode: "tracked-changes",
      idSeed: MAX_REVISION_ID - 3,
    } as const;
    const actual = appendRun({
      ...request,
      reviewer: fast.reviewer,
      anchor: fast.anchor.id,
    });
    expect(actual.applied).toHaveLength(2);
    expect(actual.skipped).toHaveLength(2);
    expect(proof.mock.results.at(0)?.value?.type).toBe("admitted");
    proof.mockReturnValue({ type: "fallback", reason: "windowShape" });
    const expected = appendRun({
      ...request,
      reviewer: slow.reviewer,
      anchor: slow.anchor.id,
    });
    expect(withoutSessionHandle(expected)).toEqual(
      withoutSessionHandle(actual),
    );
    expect(slow.reviewer.snapshot()).toStrictEqual(fast.reviewer.snapshot());
  } finally {
    proof.mockRestore();
  }
});

test("a legacy numeric alias or exhausted paragraph range declines the whole shortcut", () => {
  const content = [paragraph("0000000B", "Inserted")];
  const intent = {
    type: "insertParagraphs",
    story: documentOperations.OP_STORIES.MAIN,
    at: { type: "before", blockId: "00000001" },
    paragraphs: content,
  } satisfies documentOperations.EditorIntent;
  const op = {
    type: documentOperations.DOCUMENT_OP_TYPES.INSERT_BLOCKS,
    story: intent.story,
    at: intent.at,
    blocks: content,
  } satisfies documentOperations.DocumentOp;
  const id = {
    kind: documentOperations.ID_LEASE_KINDS.PARAGRAPH,
    id: 11,
  } as const;
  const member = { id: "insert", ops: [op], ids: [id] };
  expect(replayIds.planReplayParagraphIds(undefined, [member]).type).toBe(
    "admitted",
  );
  const lastId = documentOperations.ID_KIND_LIMITS.paragraph - 1;
  const lastParagraphId = documentOperations.formatParaId(lastId);
  const atLimit = replayIds.planReplayParagraphIds(undefined, [
    {
      ...member,
      ops: [{ ...op, blocks: [paragraph(lastParagraphId, "Last free")] }],
      ids: [{ ...id, id: lastId }],
    },
  ]);
  if (atLimit.type !== "admitted")
    throw new Error("Expected the last valid paragraph id");
  const lease = atLimit.leases.get(member.id);
  if (lease === undefined)
    throw new Error("Expected the accepted paragraph lease");
  const template = createEmptyDocument();
  const seed = paragraph(intent.at.blockId, "Anchor");
  const compile = (bodyBlocks: Paragraph[]) =>
    documentOperations.compileEditorIntent(
      {
        ...template,
        package: { ...template.package, document: { content: bodyBlocks } },
      },
      { intent, mode: { type: "editing", idLease: lease } },
    );
  expect(compile([seed]).unwrap().selection.blockId).toBe(lastParagraphId);
  const exhausted = compile([
    seed,
    paragraph(lastParagraphId, "Already occupied"),
  ]);
  if (exhausted.isOk())
    throw new Error("Expected the native paragraph allocator to refuse");
  expect(exhausted.error.reason).toBe(
    documentOperations.DOCUMENT_OP_REFUSAL_REASONS.NEEDS_NEW_IDS,
  );

  // Numeric-alias suppression during admission leaves this list empty even
  // though the native insertion had an exact, canonical paragraph ID.
  for (const ids of [
    [],
    [{ ...id, id: documentOperations.ID_KIND_LIMITS.paragraph }],
  ]) {
    expect(
      replayIds.planReplayParagraphIds(undefined, [{ ...member, ids }]),
    ).toEqual({
      type: "fallback",
      reason: "identity",
    });
  }
  expect(replayIds.planReplayParagraphIds("lease", [])).toEqual({
    type: "fallback",
    reason: "lease",
  });
});

test("a table insertion in a replay window counts one shape fallback", async () => {
  const { reviewer, anchor } = await open(await fixture());
  const proof = spyOn(replayIds, "planReplayParagraphIds");
  try {
    const result = reviewer.applyDocumentOperations(
      {
        version: 1,
        mode: "direct",
        atomic: true,
        operations: [
          {
            type: "insertBeforeBlock",
            blockId: anchor.id,
            id: "first",
            text: "First",
          },
          {
            type: "insertTable",
            blockId: anchor.id,
            id: "table",
            position: "before",
            rows: [["Term"]],
          },
          {
            type: "insertBeforeBlock",
            blockId: anchor.id,
            id: "last",
            text: "Last",
          },
        ],
      },
      { undefinedReferences: "refuse" },
    );
    expect(result.skipped).toEqual([]);
    expect(result.applied).toHaveLength(3);
    expect(proof.mock.results.map(({ value }) => value)).toEqual([
      { type: "fallback", reason: "windowShape" },
    ]);
  } finally {
    proof.mockRestore();
  }
});

const withoutCursor = () => ({
  nextLease: () => ({ type: "ordinary" as const }),
  commit: () => {},
});

test.each([12, 24, 48])(
  "reuses the occupied prefix during admission of %i paragraphs",
  async (count) => {
    const input = await fixture();
    const fast = await open(input);
    const slow = await open(input);
    const format = spyOn(documentOperations, "formatParaId");
    let admissionFormats = 0;
    const plan = replayIds.planReplayParagraphIds;
    const replay = spyOn(
      replayIds,
      "planReplayParagraphIds",
    ).mockImplementation((rejection, members) => {
      admissionFormats = format.mock.calls.length;
      return plan(rejection, members);
    });
    const start = replayIds.startParagraphAdmissionCursor;
    const seen: string[] = [];
    const cursor = spyOn(
      replayIds,
      "startParagraphAdmissionCursor",
    ).mockImplementation((options) => {
      const active = start(options);
      return {
        ...active,
        nextLease: () => {
          const lease = active.nextLease();
          seen.push(lease.type);
          return lease;
        },
      };
    });
    const compile = spyOn(documentOperations, "compileEditorIntent");
    const operations = () =>
      compile.mock.results.flatMap(({ value }) =>
        value?.isOk() ? value.value.ops : [],
      );
    const request = {
      count,
      edge: "insertBeforeBlock",
      mode: "tracked-changes",
    } as const;
    try {
      const actual = appendRun({
        ...request,
        reviewer: fast.reviewer,
        anchor: fast.anchor.id,
      });
      expect(actual.skipped).toEqual([]);
      expect(actual.applied).toHaveLength(count);
      expect(cursor).toHaveBeenCalledTimes(1);
      expect(seen).toEqual([
        "ordinary",
        ...Array.from({ length: count - 1 }, () => "bounded"),
      ]);
      const previous = admissionFormats;
      const actualOps = operations();
      format.mockClear();
      compile.mockClear();
      cursor.mockImplementation(withoutCursor);
      const ordinary = appendRun({
        ...request,
        reviewer: slow.reviewer,
        anchor: slow.anchor.id,
      });
      expect(withoutSessionHandle(ordinary)).toEqual(
        withoutSessionHandle(actual),
      );
      expect(operations()).toStrictEqual(actualOps);
      expect(slow.reviewer.snapshot()).toStrictEqual(fast.reviewer.snapshot());
      expect(previous).toBeLessThanOrEqual(5 * count + 20);
      expect(admissionFormats).toBeGreaterThanOrEqual(
        (count * (count + 1)) / 2,
      );
    } finally {
      compile.mockRestore();
      cursor.mockRestore();
      replay.mockRestore();
      format.mockRestore();
    }
  },
);

test("a refused draft leaves its selected paragraph id available for the next member", async () => {
  const input = await fixture();
  const fast = await open(input);
  const slow = await open(input);
  const ordinaryCompile = documentOperations.compileEditorIntent;
  const ordinaryReplay = replayIds.planReplayParagraphIds;
  let inAdmission = true;
  const replay = spyOn(replayIds, "planReplayParagraphIds").mockImplementation(
    (rejection, members) => {
      inAdmission = false;
      return ordinaryReplay(rejection, members);
    },
  );
  const cursor = spyOn(replayIds, "startParagraphAdmissionCursor");
  const attempts: { text: string; id: string | undefined; bounded: boolean }[] =
    [];
  const compile = spyOn(
    documentOperations,
    "compileEditorIntent",
  ).mockImplementation((document, options) => {
    const compiled = ordinaryCompile(document, options);
    if (
      !inAdmission ||
      options.intent.type !== "insertParagraphs" ||
      compiled.isErr()
    )
      return compiled;
    const leaf = options.intent.paragraphs.at(0)?.content.at(0);
    const text = leaf?.type === "run" ? leaf.content.at(0) : undefined;
    const original = compiled.value.ops.at(0);
    if (
      text?.type !== "text" ||
      original?.type !== documentOperations.DOCUMENT_OP_TYPES.INSERT_BLOCKS
    )
      return compiled;
    attempts.push({
      text: text.text,
      id: original.blocks.at(0)?.paraId,
      bounded: options.mode.idLease !== undefined,
    });
    const selected = original.blocks.at(0);
    if (text.text !== "Inserted 2" || selected === undefined) return compiled;
    // A real native paragraph collision after compilation has already
    // selected the draft's identity. Both paths still run native apply.
    return Result.ok({
      ...compiled.value,
      ops: [
        {
          ...original,
          blocks: [{ ...selected, paraId: fast.anchor.id }],
        },
      ],
    });
  });
  const request = {
    count: 4,
    edge: "insertBeforeBlock",
    mode: "tracked-changes",
    idSeed: 101,
  } as const;
  try {
    const actual = appendRun({
      ...request,
      reviewer: fast.reviewer,
      anchor: fast.anchor.id,
    });
    expect(actual.applied).toHaveLength(3);
    expect(actual.skipped).toHaveLength(1);
    expect(actual.skipped.at(0)?.message).toContain(
      "paragraph id is already used",
    );
    const refused = attempts.filter(({ text }) => text === "Inserted 2");
    expect(refused).toHaveLength(2);
    expect(refused.map(({ bounded }) => bounded)).toEqual([true, false]);
    const failedId = refused.at(0)?.id;
    expect(failedId).toBeDefined();
    expect(refused.at(1)?.id).toBe(failedId);
    expect(attempts.find(({ text }) => text === "Inserted 1")?.id).toBe(
      failedId,
    );
    cursor.mockImplementation(withoutCursor);
    inAdmission = true;
    attempts.length = 0;
    const expected = appendRun({
      ...request,
      reviewer: slow.reviewer,
      anchor: slow.anchor.id,
    });
    expect(withoutSessionHandle(expected)).toEqual(
      withoutSessionHandle(actual),
    );
    expect(slow.reviewer.snapshot()).toStrictEqual(fast.reviewer.snapshot());
    expect(attempts.filter(({ text }) => text === "Inserted 2")).toHaveLength(
      1,
    );
  } finally {
    compile.mockRestore();
    cursor.mockRestore();
    replay.mockRestore();
  }
});

test("the cursor reaches paragraph-space exhaustion and falls back with native refusal and receipts", async () => {
  const input = await fixture();
  const fast = await open(input);
  const slow = await open(input);
  const fresh = documentOperations.freshLeasedIds;
  const last = documentOperations.ID_KIND_LIMITS.paragraph - 1;
  const index = spyOn(documentOperations, "freshLeasedIds").mockImplementation(
    (options) =>
      fresh(
        options.kind === documentOperations.ID_LEASE_KINDS.PARAGRAPH &&
          options.lease === undefined
          ? // Represent an occupied legal prefix without billions of blocks.
            // The real allocator still probes the final slot and reports exhaustion.
            { ...options, from: last }
          : options,
      ),
  );
  const actualStart = replayIds.startParagraphAdmissionCursor;
  const used: string[] = [];
  const cursor = spyOn(
    replayIds,
    "startParagraphAdmissionCursor",
  ).mockImplementation((options) => {
    const active = actualStart(options);
    return {
      ...active,
      nextLease: () => {
        const lease = active.nextLease();
        used.push(lease.type === "fallback" ? lease.reason : lease.type);
        return lease;
      },
    };
  });
  const request = {
    count: 3,
    edge: "insertBeforeBlock",
    mode: "tracked-changes",
    idSeed: 101,
  } as const;
  try {
    const actual = appendRun({
      ...request,
      reviewer: fast.reviewer,
      anchor: fast.anchor.id,
    });
    expect(used).toEqual(["ordinary", "exhausted", "exhausted"]);
    expect(actual.applied).toHaveLength(1);
    expect(actual.skipped).toHaveLength(2);
    expect(actual.skipped.map(({ message }) => message)).toEqual([
      "The paragraph identity space is exhausted.",
      "The paragraph identity space is exhausted.",
    ]);
    cursor.mockImplementation(withoutCursor);
    const expected = appendRun({
      ...request,
      reviewer: slow.reviewer,
      anchor: slow.anchor.id,
    });
    expect(withoutSessionHandle(expected)).toEqual(
      withoutSessionHandle(actual),
    );
    expect(slow.reviewer.snapshot()).toStrictEqual(fast.reviewer.snapshot());
  } finally {
    cursor.mockRestore();
    index.mockRestore();
  }
});
