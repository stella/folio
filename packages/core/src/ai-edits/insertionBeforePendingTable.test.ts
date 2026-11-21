import { panic } from "better-result";
import { expect, test } from "bun:test";

import type { FolioDocumentOperation } from "../document-operations";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";
import type { HeadlessReviewerSession } from "../types/canonicalCapabilities";
import { FolioDocxReviewer } from "./headless";
import type { FolioAIEditApplyMode } from "./types";

// A paragraph put after one that stands before a tracked table takes over
// that paragraph's mark with the review records on it; accepting reads as
// the same edits made directly, rejecting gives the document back.

type Step = (ids: { left: string }) => FolioDocumentOperation;

const ALIGN: Step = ({ left }) => ({
  id: "align",
  type: "setBlockParagraphProperties",
  blockId: left,
  properties: { alignment: "center" },
});
const TABLE: Step = ({ left }) => ({
  id: "table",
  type: "insertTable",
  blockId: left,
  position: "after",
  rows: [["Term", "Value"]],
});
const ADD: Step = ({ left }) => ({
  id: "add",
  type: "insertAfterBlock",
  blockId: left,
  text: "Added words.",
});

type Case = {
  name: string;
  /** The body's paragraphs; the first is the one the steps act on. */
  body: readonly string[];
  batches: readonly Step[];
};

const CASES: readonly Case[] = [
  {
    name: "a pending inserted table",
    body: ["Left words.", "Right words."],
    batches: [TABLE, ADD],
  },
  {
    // The table takes over the paragraph's mark, which is then pending
    // inserted, and the new paragraph takes that one over.
    name: "a pending inserted table that ends the body",
    body: ["Left words."],
    batches: [TABLE, ADD],
  },
  {
    name: "a pending property change and inserted table",
    body: ["Left words.", "Right words."],
    batches: [ALIGN, TABLE, ADD],
  },
];

const SESSIONS: readonly HeadlessReviewerSession[] = ["canonical", "legacy"];

const open = async (session: HeadlessReviewerSession, body: Case["body"]) =>
  FolioDocxReviewer.fromBuffer(
    await createDocx(fromMarkdown(body.join("\n\n"))),
    { session },
  );

const leftOf = (reviewer: FolioDocxReviewer) =>
  reviewer.getContent().find(({ text }) => text === "Left words.")?.id ??
  panic("The fixture needs its first paragraph.");

type RunOptions = {
  session: HeadlessReviewerSession;
  mode: FolioAIEditApplyMode;
  testCase: Case;
};

const run = async ({
  session,
  mode,
  testCase: { body, batches },
}: RunOptions) => {
  const reviewer = await open(session, body);
  const left = leftOf(reviewer);
  for (const step of batches) {
    const result = reviewer.applyDocumentOperations(
      { version: 1, mode, operations: [step({ left })] },
      { undefinedReferences: "refuse" },
    );
    expect(result.skipped).toEqual([]);
  }
  return reviewer;
};

/** What a reader sees of the body: each block but its id and how stable it is. */
const read = (reviewer: FolioDocxReviewer) =>
  reviewer
    .getContent()
    .map(({ id: _id, idStability: _stability, ...block }) => block);

for (const session of SESSIONS)
  for (const testCase of CASES)
    test(`${session}: a paragraph after one before ${testCase.name} accepts as direct and rejects to the original`, async () => {
      const tracked = () => run({ session, mode: "tracked-changes", testCase });
      const original = read(await open(session, testCase.body));
      const direct = read(await run({ session, mode: "direct", testCase }));
      expect(
        direct.slice(0, 4).map((block) => ("text" in block ? block.text : "")),
      ).toEqual(["Left words.", "Added words.", "Term", "Value"]);
      const accepted = await tracked();
      accepted.acceptAll();
      expect(read(accepted)).toEqual(direct);
      const rejected = await tracked();
      rejected.rejectAll();
      expect(read(rejected)).toEqual(original);
      const reopened = await FolioDocxReviewer.fromBuffer(
        await (await tracked()).toBuffer(),
        { session },
      );
      reopened.rejectAll();
      expect(read(reopened)).toEqual(original);
    });
