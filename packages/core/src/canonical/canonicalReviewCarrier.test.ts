/**
 * A paragraph carrying a terminal table's review opens in a canonical
 * session, is edited as any paragraph is, and keeps its role through the
 * edit, undo and a save.
 */
import { expect, test } from "bun:test";
import { EditorState } from "prosemirror-state";

import { REVIEW_CARRIERS } from "@stll/docx-core/model";
import {
  DOCUMENT_OP_TYPES,
  INHERIT_RUN_PROPS,
  OP_STORIES,
} from "@stll/docx-core/ops";

import { parseDocx } from "../docx/parser";
import { createDocx } from "../docx/rezip";
import type { Document, Paragraph } from "../types/document";
import { readModel, type ReadonlyDocument } from "../types/documentModel";
import {
  createCanonicalSession,
  publishCanonicalProjection,
} from "./canonicalSession";

const carrierDocument = (): Document => ({
  package: {
    document: {
      content: [
        {
          type: "paragraph",
          paraId: "10000001",
          content: [{ type: "run", content: [{ type: "text", text: "Lead" }] }],
        },
        {
          type: "paragraph",
          paraId: "10000002",
          reviewCarrier: REVIEW_CARRIERS.TERMINAL_TABLE,
          content: [],
        },
      ],
    },
  },
});

const carrierOf = (document: ReadonlyDocument): Paragraph | undefined =>
  readModel(document).package.document.content.find(
    (block): block is Paragraph =>
      block.type === "paragraph" && block.reviewCarrier !== undefined,
  );

test("a review carrier paragraph projects, takes text, undoes and saves", async () => {
  const session = createCanonicalSession(carrierDocument()).unwrap();
  expect(session.projection.doc.lastChild?.attrs["reviewCarrier"]).toBe(
    REVIEW_CARRIERS.TERMINAL_TABLE,
  );
  let state = EditorState.create({ doc: session.projection.doc });
  const at = { story: OP_STORIES.MAIN, blockId: "10000002", offset: 0 };
  const prepared = session.prepareOps(
    state,
    [
      {
        type: DOCUMENT_OP_TYPES.INSERT_TEXT,
        at,
        text: "Typed",
        runProps: INHERIT_RUN_PROPS,
      },
    ],
    {
      type: "text",
      anchor: { ...at, offset: 5 },
      head: { ...at, offset: 5 },
    },
  );
  if (prepared.isErr()) throw prepared.error;
  state = publishCanonicalProjection({
    session,
    state,
    commit: prepared.value,
  }).unwrap().state;
  const typed = carrierOf(session.document);
  expect(typed?.paraId).toBe("10000002");
  expect(JSON.stringify(typed?.content)).toContain("Typed");
  const reopened = await parseDocx(
    await createDocx(readModel(session.document)),
    {
      preloadFonts: false,
    },
  );
  expect(carrierOf(reopened)?.reviewCarrier).toBe(
    REVIEW_CARRIERS.TERMINAL_TABLE,
  );
  publishCanonicalProjection({
    session,
    state,
    commit: session.prepareUndo(state).unwrap(),
  }).unwrap();
  expect(carrierOf(session.document)?.content).toEqual([]);
});
