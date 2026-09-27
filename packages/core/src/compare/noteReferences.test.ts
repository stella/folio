/**
 * Note references a comparison moves, copies or renumbers.
 *
 * Every reader shows a note reference as its reading-order marker (`[^1]`),
 * and the comparison plans on that text. Two things follow. A reference both
 * documents keep must read the same on both sides even when an added or
 * removed reference before it renumbers the revised document; and a
 * reference the comparison brings (a moved, copied or split paragraph) is
 * written as text, which has to become the reference again rather than a
 * literal `[^1]` or a bare note id in the redline.
 */

import { describe, expect, test } from "bun:test";

import { FolioDocxReviewer } from "../ai-edits/headless";
import { fromMarkdown } from "../markdown";
import { createDocx, ensureParaIds } from "../server";
import type { Endnote, Footnote, Paragraph, RunContent } from "../types/document";
import { compareDocx } from "./compare";

type Piece = string | { footnote: number } | { endnote: number };

const FOOTNOTE_IDS = [30, 10, 20, 40] as const;
const ENDNOTE_IDS = [7] as const;

const noteBody = (text: string): Paragraph => ({
  type: "paragraph",
  formatting: {},
  content: [{ type: "run", formatting: {}, content: [{ type: "text", text }] }],
});

const runContentOf = (piece: Piece): RunContent => {
  if (typeof piece === "string") return { type: "text", text: piece };
  if ("footnote" in piece) return { type: "footnoteRef", id: piece.footnote };
  return { type: "endnoteRef", id: piece.endnote };
};

/** Paragraphs of text and references; both sides carry the same notes. */
const buildDocx = async (paragraphs: readonly (readonly Piece[])[]): Promise<ArrayBuffer> => {
  const model = fromMarkdown(paragraphs.map((_, index) => `Paragraph ${index}.`).join("\n\n"));
  for (const [index, pieces] of paragraphs.entries()) {
    const paragraph = model.package.document.content[index];
    if (paragraph?.type !== "paragraph") throw new Error("fixture paragraph is missing");
    paragraph.content = [{ type: "run", formatting: {}, content: pieces.map(runContentOf) }];
  }
  model.package.footnotes = FOOTNOTE_IDS.map(
    (id): Footnote => ({
      type: "footnote",
      id,
      noteType: "normal",
      content: [noteBody(`Footnote ${id}.`)],
    }),
  );
  model.package.endnotes = ENDNOTE_IDS.map(
    (id): Endnote => ({
      type: "endnote",
      id,
      noteType: "normal",
      content: [noteBody(`Endnote ${id}.`)],
    }),
  );
  const bytes = (await ensureParaIds(new Uint8Array(await createDocx(model)))).docx;
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
};

const OPTIONS = { author: "Compare", timestamp: "2026-09-27T00:00:00.000Z" };

const resolved = async (redlined: ArrayBuffer, view: "accept" | "reject") => {
  const reviewer = await FolioDocxReviewer.fromBuffer(redlined);
  if (view === "accept") reviewer.acceptAll();
  else reviewer.rejectAll();
  return await reviewer.toBuffer();
};

const changesBetween = async (left: ArrayBuffer, right: ArrayBuffer) => {
  const result = await compareDocx(left, right, OPTIONS);
  if (result.isErr()) throw result.error;
  return result.value.changes;
};

/** Each block's text with where its references sit, which literal marker text does not have. */
const referencesOf = async (docx: ArrayBuffer) =>
  (await FolioDocxReviewer.fromBuffer(docx))
    .getContent()
    .map(({ text, structuralBoundaries }) => ({ text, structuralBoundaries }));

const BASE: readonly (readonly Piece[])[] = [
  ["Opening terms", { footnote: 30 }, " apply."],
  ["The seller delivers", { footnote: 10 }, " the goods", { endnote: 7 }, "."],
  ["Payment falls due", { footnote: 20 }, " within thirty days."],
  ["Either party may terminate", { footnote: 40 }, " on notice."],
];

const expectExactRedline = async (target: readonly (readonly Piece[])[]) => {
  const base = await buildDocx(BASE);
  const revised = await buildDocx(target);
  const compared = await compareDocx(base, revised, OPTIONS);
  if (compared.isErr()) throw compared.error;
  const { buffer, changes } = compared.value;
  expect(changes.length).toBeGreaterThan(0);
  // Accepting returns the revised document and rejecting the base, compared
  // with references read as references.
  expect(await changesBetween(revised, await resolved(buffer, "accept"))).toEqual([]);
  expect(await changesBetween(base, await resolved(buffer, "reject"))).toEqual([]);
  // A reference written back as literal `[^1]` or as its bare note id reads
  // as text, not as a reference.
  expect(await referencesOf(await resolved(buffer, "accept"))).toEqual(await referencesOf(revised));
};

describe("comparing documents whose note references move", () => {
  test("a moved paragraph carries its references", async () => {
    await expectExactRedline([BASE[0]!, BASE[2]!, BASE[3]!, BASE[1]!]);
  });

  test("a copied paragraph repeats its references, and a removed one renumbers the rest", async () => {
    await expectExactRedline([
      ["A new clause cites", { footnote: 20 }, " again."],
      BASE[0]!,
      BASE[2]!,
      ["Either party may terminate", { footnote: 40 }, " on written notice."],
    ]);
  });

  test("a split paragraph and a merged one keep their references", async () => {
    await expectExactRedline([
      [
        "Opening terms",
        { footnote: 30 },
        " apply. The seller delivers",
        { footnote: 10 },
        " the goods",
        { endnote: 7 },
        ".",
      ],
      ["Payment falls due", { footnote: 20 }, " on delivery."],
      ["Within thirty days of it."],
      BASE[3]!,
    ]);
  });

  test("a paragraph the revision adds cites a note the base already has", async () => {
    await expectExactRedline([
      BASE[0]!,
      BASE[1]!,
      ["A new schedule, see", { endnote: 7 }, " and", { footnote: 30 }, "."],
      BASE[2]!,
      BASE[3]!,
    ]);
  });
});
