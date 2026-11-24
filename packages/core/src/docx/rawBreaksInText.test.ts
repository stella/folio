/**
 * Line breaks and tabs written as characters inside `w:t` or `w:delText`
 * read as a word processor reads them: a CR LF pair is one space, each other
 * CR or LF is one space, none collapsing into its neighbours, and a tab
 * character is a tab element. Both engines open such a package, and a save
 * keeps the spaces.
 */
import { describe, expect, test } from "bun:test";

import { buildRawPackage, SECTION } from "../__tests__/documentShapes";
import { FolioDocxReviewer } from "../ai-edits/headless";
import { parseDocx } from "./parser";

const SESSIONS = ["canonical"] as const;

/** Each raw spelling inside `w:t`, and the text it reads as. */
const SHAPES = [
  { name: "a line feed", raw: "Alpha&#10;Beta", text: "Alpha Beta" },
  { name: "a carriage return", raw: "Alpha&#13;Beta", text: "Alpha Beta" },
  { name: "a literal CR LF", raw: "Alpha\r\nBeta", text: "Alpha Beta" },
  { name: "a CR LF pair", raw: "Alpha&#13;&#10;Beta", text: "Alpha Beta" },
  { name: "an LF CR pair", raw: "Alpha&#10;&#13;Beta", text: "Alpha  Beta" },
  { name: "two line feeds", raw: "Alpha&#10;&#10;Beta", text: "Alpha  Beta" },
  {
    name: "two carriage returns",
    raw: "Alpha&#13;&#13;Beta",
    text: "Alpha  Beta",
  },
  {
    name: "a space and a line feed",
    raw: "Alpha &#10;Beta",
    text: "Alpha  Beta",
  },
  { name: "a leading line feed", raw: "&#10;Alpha", text: " Alpha" },
  { name: "a trailing line feed", raw: "Alpha&#10;", text: "Alpha " },
] as const;

const packageWith = async (inner: string): Promise<ArrayBuffer> => {
  const bytes = await buildRawPackage({
    body: `<w:p>${inner}</w:p>${SECTION()}`,
    styles: null,
  });
  return bytes.slice().buffer;
};

const textRun = (raw: string) =>
  `<w:r><w:t xml:space="preserve">${raw}</w:t></w:r>`;

const runContentTypes = async (buffer: ArrayBuffer) => {
  const document = await parseDocx(buffer, { preloadFonts: false });
  const paragraph = document.package.document.content.at(0);
  if (paragraph?.type !== "paragraph") throw new Error("Expected a paragraph");
  return paragraph.content.flatMap((item) =>
    item.type === "run"
      ? item.content.map((content) =>
          content.type === "text" ? content.text : content.type,
        )
      : [],
  );
};

describe("raw line breaks inside text", () => {
  for (const { name, raw, text } of SHAPES) {
    test(`${name} reads as its spaces`, async () => {
      expect(await runContentTypes(await packageWith(textRun(raw)))).toEqual([
        text,
      ]);
    });

    for (const session of SESSIONS)
      test(`${name} opens and saves its spaces (${session})`, async () => {
        const reviewer = await FolioDocxReviewer.fromBuffer(
          await packageWith(textRun(raw)),
          { session },
        );
        expect(reviewer.snapshot().blocks.at(0)?.text).toBe(text);
        const reopened = await FolioDocxReviewer.fromBuffer(
          await reviewer.toBuffer(),
          { session },
        );
        expect(reopened.snapshot().blocks.at(0)?.text).toBe(text);
      });
  }

  test("a tab character becomes a tab element between the text around it", async () => {
    expect(
      await runContentTypes(await packageWith(textRun("Alpha\tBeta\t"))),
    ).toEqual(["Alpha", "tab", "Beta", "tab"]);
  });

  for (const session of SESSIONS)
    test(`tracked text with raw breaks resolves as plain text (${session})`, async () => {
      const buffer = await packageWith(
        textRun("Gamma ") +
          '<w:ins w:id="1" w:author="A" w:date="2026-01-01T00:00:00Z">' +
          textRun("Delta&#10;Epsilon") +
          "</w:ins>" +
          '<w:del w:id="2" w:author="A" w:date="2026-01-01T00:00:00Z">' +
          '<w:r><w:delText xml:space="preserve">&#13;&#10;Zeta</w:delText></w:r>' +
          "</w:del>",
      );
      const accepting = await FolioDocxReviewer.fromBuffer(buffer, {
        session,
      });
      accepting.acceptAll();
      expect(accepting.snapshot().blocks.at(0)?.text).toBe(
        "Gamma Delta Epsilon",
      );
      const rejecting = await FolioDocxReviewer.fromBuffer(buffer, {
        session,
      });
      rejecting.rejectAll();
      expect(rejecting.snapshot().blocks.at(0)?.text).toBe("Gamma  Zeta");
    });
});
