import { panic } from "better-result";
import { expect, test } from "bun:test";
import { EditorState } from "prosemirror-state";

import {
  DOCUMENT_OP_TYPES,
  type InsertBlocksOp,
  OP_STORIES,
} from "@stll/docx-core/ops";

import { canonicalOperationStagingAccess } from "../internal/canonicalOperationStaging";
import {
  markupViewSources,
  markupViewStateSources,
} from "../prosemirror/conversion/markupViewSource";
import { readModel } from "../types/documentModel";
import {
  CANONICAL_ADMISSION,
  createCanonicalSession,
  publishCanonicalProjection,
} from "./canonicalSession";

test("insertion staging owns the verified operation and rejects foreign or stale results", () => {
  const session = createCanonicalSession({
    package: {
      document: {
        content: [{ type: "paragraph", paraId: "00000001", content: [] }],
      },
    },
  }).unwrap();
  const source = readModel(session.document);
  const blockId =
    session.projection.firstParagraph?.blockId ??
    panic("The fixture has no paragraph anchor.");
  const state = EditorState.create({ doc: session.projection.doc });
  const conversion = markupViewSources.get(state.doc)?.options;
  expect(conversion?.blockCache).toBeDefined();
  const selection = session.projection.selectionAt(state).unwrap();
  const text = { type: "text" as const, text: "Verified." };
  const block = {
    type: "paragraph",
    paraId: blockId,
    content: [{ type: "run", content: [text] }],
  } satisfies InsertBlocksOp["blocks"][number];
  const op = {
    type: DOCUMENT_OP_TYPES.INSERT_BLOCKS,
    story: OP_STORIES.MAIN,
    at: { type: "after", blockId },
    blocks: [block],
  } satisfies InsertBlocksOp;
  expect(() =>
    Reflect.apply(session.beginVerifiedOperationStaging, session, [Symbol()]),
  ).toThrow("Only the public-operation compiler");
  const staging = session.beginVerifiedOperationStaging(
    canonicalOperationStagingAccess,
    CANONICAL_ADMISSION.edit,
  );
  expect(staging.apply(source, op).isErr()).toBe(true);
  block.paraId = "001ABCDE";
  const verified = staging.apply(source, op).unwrap();
  expect(staging.append({ ...verified })).toBe(false);
  block.paraId = "001ABCDF";
  text.text = "Mutated.";
  expect(staging.append(verified)).toBe(true);
  const commit = staging.prepare(state, selection).unwrap();
  expect(commit.ops).toEqual([
    {
      ...op,
      blocks: [
        {
          ...block,
          paraId: "001ABCDE",
          content: [
            { type: "run", content: [{ type: "text", text: "Verified." }] },
          ],
        },
      ],
    },
  ]);
  expect(commit.touched.inserted).toEqual(["001ABCDE"]);
  const published = publishCanonicalProjection({
    session,
    state,
    commit,
  }).unwrap();
  const publishedConversion = markupViewStateSources.get(
    published.state,
  )?.options;
  expect(publishedConversion?.blockCache === conversion?.blockCache).toBe(true);
  expect(publishedConversion?.textBoxGroupSalt).toBe(
    conversion?.textBoxGroupSalt,
  );
  const stale = staging.prepare(state, selection);
  if (stale.isOk()) panic("The previous version's staging should be stale.");
  expect(stale.error.message).toContain("stale canonical document");
});
