import { describe, expect, test } from "bun:test";
import { readFileSync } from "node:fs";
import path from "node:path";

import {
  FolioDocxReviewer,
  parseFolioDocumentOperationBatch,
  type FolioDocumentOperationBatch,
  type FolioDocumentOperationResult,
} from "@stll/folio-core/server";

import { createEditorRefBridge, type FolioAgentEditorRefLike } from "./bridges/editor-ref";
import { createReviewerBridge } from "./bridges/reviewer";

const AUTHOR = "Conformance";
const SOURCE_PATH = path.join(
  import.meta.dir,
  "../../core/src/docx/__tests__/__fixtures__/corpus/authored-empty-paragraph.docx",
);
const OPERATION_FIXTURE_PATH = path.join(
  import.meta.dir,
  "../../../tests/operations/tracked-replace.v1.json",
);

const readArrayBuffer = (filePath: string): ArrayBuffer => {
  const bytes = readFileSync(filePath);
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
};

const readOperationFixture = (): FolioDocumentOperationBatch =>
  parseFolioDocumentOperationBatch(JSON.parse(readFileSync(OPERATION_FIXTURE_PATH, "utf8")));

type ConformanceSurface = {
  name: string;
  reviewer: FolioDocxReviewer;
  apply(batch: FolioDocumentOperationBatch): FolioDocumentOperationResult;
};

const createEditorRef = (reviewer: FolioDocxReviewer): FolioAgentEditorRefLike => ({
  createAIEditSnapshot: () => reviewer.snapshot(),
  applyAIEditOperations: ({ operations, mode, author }) =>
    reviewer.applyOperations(operations, { mode, author }),
  applyDocumentOperations: ({ batch }) => reviewer.applyDocumentOperations(batch),
  scrollToBlock: () => false,
  getTotalPages: () => 1,
  getTrackedChanges: () => reviewer.getChanges(),
});

const createSurfaces = async (): Promise<ConformanceSurface[]> => {
  const source = readArrayBuffer(SOURCE_PATH);
  const headlessReviewer = await FolioDocxReviewer.fromBuffer(source, { author: AUTHOR });
  const reviewerBridgeReviewer = await FolioDocxReviewer.fromBuffer(source, { author: AUTHOR });
  const editorBridgeReviewer = await FolioDocxReviewer.fromBuffer(source, { author: AUTHOR });
  const reviewerBridge = createReviewerBridge(reviewerBridgeReviewer);
  const editorBridge = createEditorRefBridge({
    ref: createEditorRef(editorBridgeReviewer),
    author: AUTHOR,
    getComments: () => [],
    setComments: () => {},
  });

  return [
    {
      name: "headless",
      reviewer: headlessReviewer,
      apply: (batch) => headlessReviewer.applyDocumentOperations(batch),
    },
    {
      name: "reviewer bridge",
      reviewer: reviewerBridgeReviewer,
      apply: (batch) => reviewerBridge.applyDocumentOperations(batch),
    },
    {
      name: "editor-ref bridge",
      reviewer: editorBridgeReviewer,
      apply: (batch) => editorBridge.applyDocumentOperations(batch),
    },
  ];
};

const normalizeResult = ({ version, applied, skipped }: FolioDocumentOperationResult) => ({
  version,
  applied: applied.map(({ id }) => ({ id })),
  skipped,
});

const normalizeChanges = (reviewer: FolioDocxReviewer) =>
  reviewer.getChanges().map(({ type, author, text, blockId }) => ({
    type,
    author,
    text,
    blockId,
  }));

describe("document operation cross-surface conformance", () => {
  test("produces equivalent results and saved semantics", async () => {
    const batch = readOperationFixture();
    const outputs = [];

    for (const surface of await createSurfaces()) {
      const result = surface.apply(batch);
      const savedReviewer = await FolioDocxReviewer.fromBuffer(await surface.reviewer.toBuffer());
      outputs.push({
        name: surface.name,
        result: normalizeResult(result),
        content: surface.reviewer.getContentAsText(),
        changes: normalizeChanges(surface.reviewer),
        savedContent: savedReviewer.getContentAsText(),
        savedChanges: normalizeChanges(savedReviewer),
      });
    }

    const expected = {
      result: {
        version: 1,
        applied: [{ id: "replace-heading" }],
        skipped: [],
      },
      content: "[0304003A] Intro paragraph.\n[32560014] Trailing paragraph.",
      changes: [
        {
          type: "deletion",
          author: AUTHOR,
          text: "Heading",
          blockId: "0304003A",
        },
        {
          type: "insertion",
          author: AUTHOR,
          text: "Intro",
          blockId: "0304003A",
        },
      ],
    };

    for (const output of outputs) {
      expect(output.result, output.name).toEqual(expected.result);
      expect(output.content, output.name).toBe(expected.content);
      expect(output.changes, output.name).toEqual(expected.changes);
      expect(output.savedContent, output.name).toBe(expected.content);
      expect(output.savedChanges, output.name).toEqual(expected.changes);
    }
  });
});
