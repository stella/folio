import { expect, test } from "bun:test";

import { relationshipIdOf } from "@stll/docx-core/model";
import { ZipArchive } from "@stll/docx-core/zip";

import { FolioDocxReviewer } from "../ai-edits/headless";
import { createDocx } from "../docx/rezip";
import type { HeadlessReviewerSession } from "../types/canonicalCapabilities";
import type { ParagraphContent } from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";
import { compareDocx } from "./compare";

const SESSIONS: readonly HeadlessReviewerSession[] = ["canonical"];
const OPTIONS = {
  author: "Picture reviewer",
  timestamp: "2026-09-13T00:00:00.000Z",
} as const;
const PICTURE =
  "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=";

const text = (value: string): ParagraphContent => ({
  type: "run",
  content: [{ type: "text", text: value }],
});

const picture: ParagraphContent = {
  type: "run",
  content: [
    {
      type: "drawing",
      image: {
        type: "image",
        rId: relationshipIdOf("rId_img_picture"),
        src: PICTURE,
        filename: "picture.png",
        size: { width: 914_400, height: 914_400 },
        wrap: { type: "inline" },
      },
    },
  ],
};

const docx = (content: readonly ParagraphContent[]) => {
  const document = createEmptyDocument();
  document.package.document.content = [
    { type: "paragraph", paraId: "00000001", content: [...content] },
  ];
  return createDocx(document);
};

/** The ids of a saved document's drawings, in document order. */
const savedDrawingIds = async (docxBuffer: ArrayBuffer): Promise<string[]> => {
  const zip = await ZipArchive.read(docxBuffer);
  const xml = await zip.entry("word/document.xml")?.text();
  return [...(xml ?? "").matchAll(/<wp:docPr\b[^>]*?\sid="(?<id>\d+)"/gu)].map(
    (match) => match.groups?.["id"] ?? "",
  );
};

// The moved picture stands twice in the redline, marked deleted where it was
// and inserted where it goes, and a package numbers each drawing once: in the
// model's own fields and in the markup a saved picture keeps.
test.each(SESSIONS)(
  "%s: a picture a comparison moves takes a drawing id of its own",
  async (session) => {
    const other = structuredClone(picture);
    const base = await docx([
      text("Alpha "),
      picture,
      text("Beta "),
      other,
      text("Gamma"),
    ]);
    const target = await docx([
      text("Alpha Beta "),
      picture,
      text("Gamma"),
      other,
    ]);
    for (const side of [base, target])
      expect(new Set(await savedDrawingIds(side)).size).toBe(2);
    const compared = await compareDocx(base, target, { ...OPTIONS, session });
    if (compared.isErr()) throw compared.error;
    expect(compared.value.verification).toEqual({ status: "verified" });
    const ids = await savedDrawingIds(compared.value.buffer);
    expect(ids.length).toBeGreaterThan(2);
    expect(new Set(ids).size).toBe(ids.length);
  },
);

// A comparison copies a moved picture into a media part of its own; the
// resolved document holds the same picture, so comparing it with the side it
// resolves to reports nothing.
test.each(SESSIONS)(
  "%s: a picture a comparison moves resolves to the same picture",
  async (session) => {
    const base = await docx([text("Alpha "), picture, text("Beta")]);
    const target = await docx([text("Alpha Beta"), picture]);
    const compared = await compareDocx(base, target, { ...OPTIONS, session });
    if (compared.isErr()) throw compared.error;
    expect(compared.value.verification).toEqual({ status: "verified" });
    for (const [mode, side] of [
      ["accept", target],
      ["reject", base],
    ] as const) {
      const reviewer = await FolioDocxReviewer.fromBuffer(
        compared.value.buffer,
        { session },
      );
      if (mode === "accept") reviewer.acceptAll();
      else reviewer.rejectAll();
      const again = await compareDocx(side, await reviewer.toBuffer(), {
        ...OPTIONS,
        session,
      });
      if (again.isErr()) throw again.error;
      expect({ mode, changes: again.value.changes }).toEqual({
        mode,
        changes: [],
      });
    }
  },
);
