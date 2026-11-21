import { expect, spyOn, test } from "bun:test";
import { EditorState } from "prosemirror-state";

import { type IdLease, normalizeForOps } from "@stll/docx-core/ops";
import * as documentOperations from "@stll/docx-core/ops";

import { createFolioAIEditSnapshot } from "../ai-edits/snapshot";
import { CanonicalPublicOperations } from "../canonical/canonicalPublicOperations";
import {
  createCanonicalSession,
  publishCanonicalProjection,
} from "../canonical/canonicalSession";
import type {
  FolioDocumentOperation,
  FolioDocumentOperationMode,
} from "../document-operations";
import { readModel } from "../types/documentModel";

const source = normalizeForOps({
  package: {
    document: {
      content: [
        {
          type: "paragraph",
          paraId: "10000001",
          content: [{ type: "run", content: [{ type: "text", text: "Left" }] }],
        },
        {
          type: "paragraph",
          paraId: "10000002",
          content: [
            { type: "run", content: [{ type: "text", text: "Right" }] },
          ],
        },
      ],
    },
  },
});

const setup = (
  lease?: IdLease,
  mode: FolioDocumentOperationMode = "direct",
) => {
  const session = createCanonicalSession(
    source,
    undefined,
    lease === undefined ? {} : { idLease: lease },
  ).unwrap();
  let state = EditorState.create({ doc: session.projection.doc });
  const executor = new CanonicalPublicOperations({
    session,
    getState: () => state,
    publish: (commit) => {
      const published = publishCanonicalProjection({ state, session, commit });
      if (published.isErr()) return false;
      state = published.value.state;
      return true;
    },
  });
  const apply = (operations: readonly FolioDocumentOperation[]) =>
    executor.apply({
      snapshot: createFolioAIEditSnapshot(state.doc),
      batch: { version: 1, mode, operations },
      undefinedReferences: "refuse",
      revisionStamp: { idSeed: 1, date: "2026-01-02T03:04:05Z" },
    });
  return { session, apply, state: () => state };
};

test("gap planning preserves scarce normal-order admission and leaves discarded preparations unpublished", () => {
  const lease = {
    blocks: [{ kind: "paragraph", start: 0x20000001, end: 0x20000003 }],
  } as const satisfies IdLease;
  const editor = setup(lease);
  const original = structuredClone(source);
  const result = editor.apply([
    { id: "two", type: "insertAfterBlock", blockId: "10000001", text: "Y\nZ" },
    { id: "one", type: "insertBeforeBlock", blockId: "10000002", text: "X" },
  ]);
  expect(result.applied.map(({ id }) => id)).toEqual(["one"]);
  expect(result.skipped.map(({ id }) => id)).toEqual(["two"]);
  expect(editor.session.version).toBe(1);
  expect([...editor.session.reservedBlockIds].toSorted()).toEqual([
    "10000001",
    "10000002",
    "20000001",
  ]);
  expect(source).toStrictEqual(original);
  expect(lease.blocks).toEqual([
    { kind: "paragraph", start: 0x20000001, end: 0x20000003 },
  ]);
});

test.each([
  ["direct", 4],
  ["direct", 8],
  ["tracked-changes", 4],
  ["tracked-changes", 8],
] as const)(
  "leased admission avoids a package census per reordered insertion (%s, %i)",
  (mode, count) => {
    const first = 0x20000001;
    const lease = {
      blocks: [
        { kind: "paragraph", start: first, end: first + count },
        { kind: "revision", start: 1, end: 1 + 4 * count },
      ],
    } satisfies IdLease;
    const editor = setup(lease, mode);
    const operations = Array.from(
      { length: count },
      (_, index) =>
        ({
          id: `insert-${index}`,
          type: "insertBeforeBlock",
          blockId: "10000002",
          text: `Added ${index}`,
        }) satisfies FolioDocumentOperation,
    );
    const census = spyOn(documentOperations, "leasedIdCensus");
    try {
      const result = editor.apply(operations);
      expect(result.skipped).toEqual([]);
      expect(result.applied.map(({ id }) => id)).toEqual(
        operations.map(({ id }) => id),
      );
      expect(census).toHaveBeenCalledTimes(1);
      if (mode === "tracked-changes")
        expect(
          result.applied.every(({ revisionId }) => revisionId !== undefined),
        ).toBe(true);
      expect(editor.session.reservedBlockIds).toEqual(
        new Set([
          "10000001",
          "10000002",
          ...Array.from({ length: count }, (_, offset) =>
            documentOperations.formatParaId(first + offset),
          ),
        ]),
      );
    } finally {
      census.mockRestore();
    }
  },
);

test("reordered insertions retain normal-order paragraph and attached-comment allocation ownership", () => {
  const beforeRight = {
    id: "second",
    type: "insertBeforeBlock",
    blockId: "10000002",
    text: "Second",
    comment: { text: "Second comment" },
  } as const satisfies FolioDocumentOperation;
  const afterLeft = {
    id: "first",
    type: "insertAfterBlock",
    blockId: "10000001",
    text: "First",
    comment: { text: "First comment" },
  } as const satisfies FolioDocumentOperation;
  const normal = setup();
  const normalSecond = normal.apply([beforeRight]);
  const normalFirst = normal.apply([afterLeft]);
  expect(normalSecond.skipped).toEqual([]);
  expect(normalFirst.skipped).toEqual([]);
  const planned = setup();
  const result = planned.apply([afterLeft, beforeRight]);
  expect(result.skipped).toEqual([]);
  expect(result.applied).toEqual([
    ...normalFirst.applied,
    ...normalSecond.applied,
  ]);
  const paragraphs = (editor: ReturnType<typeof setup>) =>
    readModel(editor.session.document)
      .package.document.content.filter((block) => block.type === "paragraph")
      .map(({ paraId, content }) => ({ paraId, content }));
  expect(paragraphs(planned)).toStrictEqual(paragraphs(normal));
  expect(
    readModel(planned.session.document).package.document.comments,
  ).toStrictEqual(readModel(normal.session.document).package.document.comments);
  expect(planned.session.version).toBe(1);
});
