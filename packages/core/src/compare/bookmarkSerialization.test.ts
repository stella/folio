import { describe, expect, test } from "bun:test";
import JSZip from "jszip";

import { FolioDocxReviewer } from "../ai-edits/headless";
import { buildBodySequenceDocx } from "./__fixtures__/body-sequence";
import { compareDocx } from "./compare";

const documentPartOf = async (buffer: ArrayBuffer): Promise<string> =>
  (await (await JSZip.loadAsync(buffer)).file("word/document.xml")?.async("string")) ?? "";

const withTrackedBookmark = async (type: "ins" | "del", text: string): Promise<ArrayBuffer> => {
  const buffer = await buildBodySequenceDocx([{ kind: "paragraph", text: "Before after." }]);
  const zip = await JSZip.loadAsync(buffer);
  const xml = await documentPartOf(buffer);
  const textElement = type === "del" ? "delText" : "t";
  const tracked =
    `<w:${type} w:id="7" w:author="Reviewer" w:date="2026-09-06T12:00:00Z">` +
    '<w:bookmarkStart w:id="9" w:name="TrackedTerm"/>' +
    `<w:r><w:${textElement}>${text}</w:${textElement}></w:r>` +
    '<w:bookmarkEnd w:id="9"/>' +
    `</w:${type}>`;
  zip.file(
    "word/document.xml",
    xml.replace(
      '<w:r><w:t xml:space="preserve">Before after.</w:t></w:r>',
      `<w:r><w:t xml:space="preserve">Before </w:t></w:r>${tracked}<w:r><w:t>after.</w:t></w:r>`,
    ),
  );
  return await zip.generateAsync({ type: "arraybuffer" });
};

describe("tracked bookmark serialization", () => {
  test("compare removes a deleted bookmark with its accepted revision", async () => {
    const base = await withTrackedBookmark("del", "deleted ");
    const target = await buildBodySequenceDocx([{ kind: "paragraph", text: "Before after." }]);
    const result = await compareDocx(base, target, {
      author: "compare",
      timestamp: "2026-09-06T12:00:00.000Z",
    });
    if (result.isErr()) {
      throw result.error;
    }

    expect(result.value.verification).toEqual({ status: "verified" });
    const xml = await documentPartOf(result.value.buffer);
    expect(xml).not.toContain("bookmarkStart");
    expect(xml).not.toContain("bookmarkEnd");
    expect(xml).not.toContain("deleted");
  });

  test("rejecting an inserted range removes its bookmark boundaries", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(
      await withTrackedBookmark("ins", "inserted "),
    );
    const story = reviewer.listStories().at(0)?.handle;
    if (!story) {
      throw new Error("Expected document story");
    }
    reviewer.resolveReviewedStory({ story, view: "original" });

    const xml = await documentPartOf(await reviewer.toBuffer());
    expect(xml).not.toContain("bookmarkStart");
    expect(xml).not.toContain("bookmarkEnd");
    expect(xml).not.toContain("inserted");
  });
});
