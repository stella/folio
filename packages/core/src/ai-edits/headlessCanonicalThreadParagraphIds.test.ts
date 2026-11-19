/**
 * A comment the package wrote without paragraph ids gets the id its thread is
 * linked by when a reply or a resolution threads it on the canonical session,
 * in the model as in the saved file: the save mints none of its own.
 */
import { expect, test } from "bun:test";

import { createDocx } from "../docx/rezip";
import { paragraph, run } from "../docx/server/build";
import { fromMarkdown } from "../markdown/fromMarkdown";
import type { Document } from "../types/document";
import { FolioDocxReviewer } from "./headless";

const source = (): Document => {
  const document = fromMarkdown("Placeholder.");
  document.package.document.content = [
    paragraph([
      { type: "commentRangeStart", id: 41 },
      run("Commented words."),
      { type: "commentRangeEnd", id: 41 },
      { type: "commentReference", id: 41 },
    ]),
  ];
  document.package.document.comments = [
    {
      id: 41,
      author: "Earlier Reviewer",
      content: [paragraph("Check the words.")],
    },
  ];
  return document;
};

const threadIds = (document: Document) =>
  (document.package.document.comments ?? []).map(({ id, content }) => [
    id,
    content?.at(-1)?.paraId,
  ]);

const open = async (bytes: ArrayBuffer) =>
  FolioDocxReviewer.fromBuffer(bytes, {
    author: "Editor",
    session: "canonical",
  });

test.each(["reply", "resolve"] as const)(
  "a %s gives an id-less comment the id the save links it by",
  async (command) => {
    const reviewer = await open(await createDocx(source()));
    expect(threadIds(reviewer.toDocument())).toEqual([[41, undefined]]);
    if (command === "reply")
      expect(reviewer.replyTo(41, { text: "Done." })).not.toBeNull();
    else expect(reviewer.resolveComment("41")).toBe(true);
    const live = threadIds(reviewer.toDocument());
    expect(live.every(([, paraId]) => typeof paraId === "string")).toBe(true);
    const reopened = await open(await reviewer.toBuffer());
    expect(threadIds(reopened.toDocument())).toEqual(live);
  },
);
