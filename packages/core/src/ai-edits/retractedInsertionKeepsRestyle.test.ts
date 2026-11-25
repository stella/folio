/**
 * Deleting the paragraphs a reviewer put after a story's last paragraph,
 * which the reviewer restyled too, keeps the restyle: applied directly, and
 * as tracked changes accepted, which read the same once saved and reopened;
 * rejected, the tracked document reads as it was.
 *
 * The last paragraph is a list item or a plain one. A first batch puts a
 * paragraph after it and gives it a heading style, a centred alignment or
 * both, in either order; a second batch deletes the paragraph the first put
 * there. Every case runs.
 */

import { expect, test } from "bun:test";

import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
} from "../document-operations";
import { ensureParaIds } from "../docx/ensureParaIds";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";
import { FolioDocxReviewer } from "./headless";

type Mode = "direct" | "tracked-changes";
type Case = {
  last: "listItem" | "paragraph";
  restyle: "heading" | "centered" | "both";
  restyleFirst: boolean;
};

const OPENING = "Opening clause.";
const LAST = "Last clause.";
const INSERTED = "Inserted clause";

const open = async ({ last }: Case): Promise<FolioDocxReviewer> => {
  const texts = [
    `# ${OPENING}`,
    "1. Listed clause.",
    last === "listItem" ? `1. ${LAST}` : LAST,
  ];
  const bytes = (
    await ensureParaIds(await createDocx(fromMarkdown(texts.join("\n\n"))))
  ).docx;
  return FolioDocxReviewer.fromBuffer(bytes, { author: "Reviewer" });
};

const apply = (
  reviewer: FolioDocxReviewer,
  mode: Mode,
  operations: FolioDocumentOperation[],
): string[] =>
  reviewer
    .applyDocumentOperations(
      {
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        mode,
        atomic: false,
        operations,
      },
      { undefinedReferences: "refuse" },
    )
    .applied.map(({ id }) => id)
    .toSorted();

const read = (reviewer: FolioDocxReviewer) =>
  reviewer.getContent().map(({ kind, text, styleId, directAlignment }) => ({
    kind,
    text,
    styleId: styleId ?? null,
    alignment: directAlignment ?? null,
  }));

const reopened = async (reviewer: FolioDocxReviewer) =>
  FolioDocxReviewer.fromBuffer(await reviewer.toBuffer(), {
    author: "Reviewer",
  });

const replay = async (setup: Case, mode: Mode) => {
  const reviewer = await open(setup);
  const original = read(reviewer);
  const last = reviewer.getContent().find(({ text }) => text === LAST);
  if (last === undefined) throw new Error("Missing last paragraph.");
  const insertion: FolioDocumentOperation = {
    id: "insert",
    type: "insertAfterBlock",
    blockId: last.id,
    text: `${INSERTED}.`,
  };
  const restyle: FolioDocumentOperation = {
    id: "restyle",
    type: "setBlockParagraphProperties",
    blockId: last.id,
    properties: {
      ...(setup.restyle === "centered" ? {} : { styleId: "Heading2" }),
      ...(setup.restyle === "heading" ? {} : { alignment: "center" }),
    },
  };
  const first = apply(
    reviewer,
    mode,
    setup.restyleFirst ? [restyle, insertion] : [insertion, restyle],
  );
  const second = apply(
    reviewer,
    mode,
    reviewer
      .getContent()
      .filter(({ text }) => text.startsWith(INSERTED))
      .map(({ id }) => ({ id: "delete", type: "deleteBlock", blockId: id })),
  );
  const rejected = await reopened(reviewer);
  rejected.rejectAll();
  reviewer.acceptAll();
  return {
    applied: [first, second],
    accepted: read(await reopened(reviewer)),
    rejected: read(rejected),
    original,
  };
};

const CASES: Case[] = (["listItem", "paragraph"] as const).flatMap((last) =>
  (["heading", "centered", "both"] as const).flatMap((restyle) =>
    [true, false].map((restyleFirst) => ({ last, restyle, restyleFirst })),
  ),
);

// #1350 flow b03f9f4bc6b1ab1d is the list item restyled as a heading.
test.each(CASES)(
  "deleting the paragraph put after a restyled last paragraph keeps the restyle in every mode: %o",
  async (setup) => {
    const direct = await replay(setup, "direct");
    const tracked = await replay(setup, "tracked-changes");
    expect(tracked.applied).toEqual(direct.applied);
    expect(tracked.applied).toEqual([["insert", "restyle"], ["delete"]]);
    expect(tracked.accepted).toEqual(direct.accepted);
    expect(tracked.rejected).toEqual(tracked.original);
  },
);
