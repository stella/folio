import { expect, test } from "bun:test";
import { InvalidOoxmlNumericIdError } from "@stll/docx-core";
import JSZip from "jszip";

import { createEmptyDocument } from "../utils/createDocument";
import { parseDocx } from "./parser";
import { createDocx, createEmptyDocx, repackDocx } from "./rezip";

test("the DOCX writer rejects a synthetic timestamp-sized comment id", async () => {
  const doc = createEmptyDocument();
  doc.package.document.comments = [{ id: 1_784_212_345_678, author: "Reviewer", content: [] }];
  await expect(createDocx(doc)).rejects.toBeInstanceOf(InvalidOoxmlNumericIdError);
});

test("the DOCX writer rejects invalid ids injected into untouched note parts after import", async () => {
  const doc = await parseDocx(await createEmptyDocx());
  const zip = await JSZip.loadAsync(await createEmptyDocx());
  zip.file(
    "word/footnotes.xml",
    '<x:footnotes xmlns:x="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><x:footnote x:type="separator" x:id="2147483648"><x:p/></x:footnote></x:footnotes>',
  );
  const source = await zip.generateAsync({ type: "arraybuffer" });
  doc.originalBuffer = source;
  await expect(repackDocx(doc)).rejects.toBeInstanceOf(InvalidOoxmlNumericIdError);
});

test("valid high comment and bookmark ids survive save and reopen", async () => {
  const doc = createEmptyDocument();
  const commentId = 2_147_483_646;
  const bookmarkId = 2_147_483_647;
  doc.package.document.comments = [{ id: commentId, author: "Reviewer", content: [] }];
  doc.package.document.content = [
    {
      type: "paragraph",
      content: [
        { type: "bookmarkStart", id: bookmarkId, name: "marker" },
        { type: "commentRangeStart", id: commentId },
        { type: "run", content: [{ type: "text", text: "Anchor" }] },
        { type: "commentRangeEnd", id: commentId },
        { type: "bookmarkEnd", id: bookmarkId },
      ],
    },
  ];
  const reopened = await parseDocx(await createDocx(doc));
  expect(reopened.package.document.comments?.at(0)?.id).toBe(commentId);
  const block = reopened.package.document.content.at(0);
  expect(block?.type).toBe("paragraph");
  if (block?.type !== "paragraph") return;
  expect(block.content.find((content) => content.type === "bookmarkStart")).toMatchObject({
    id: bookmarkId,
  });
  const twice = await parseDocx(await repackDocx(reopened));
  expect(twice.package.document.comments?.at(0)?.id).toBe(commentId);
});
