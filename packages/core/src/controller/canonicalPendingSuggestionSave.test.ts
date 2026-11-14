import { describe, expect, test } from "bun:test";

import { ZipArchive } from "@stll/docx-core/zip";

import { createCanonicalSession } from "../canonical/canonicalSession";
import { serializeCanonicalSave } from "../docx/canonicalSave";
import { prepareCanonicalDocxInput } from "../docx/canonicalSessionInput";
import { parseDocx } from "../docx/parser";
import { createDocx } from "../docx/rezip";
import type { Document, ParagraphContent } from "../types/document";

const run = (text: string): ParagraphContent => ({
  type: "run",
  content: [{ type: "text", text }],
});
const info = (id: number, suggestionId?: string) => ({
  id,
  author: "Reviewer",
  date: "2026-01-01T00:00:00Z",
  ...(suggestionId === undefined ? {} : { suggestionId }),
});

const source: Document = {
  package: {
    document: {
      content: [
        { type: "paragraph", paraId: "10000000", content: [run("Hello")] },
      ],
    },
  },
};

/** The saved body text and its tracked change elements. */
const savedBody = async (content: ParagraphContent[]) => {
  const session = createCanonicalSession(
    await parseDocx(
      (await prepareCanonicalDocxInput(await createDocx(source))).unwrap(),
      { preloadFonts: false, detectVariables: false },
    ),
  ).unwrap();
  const snapshot = session.captureSaveSnapshot();
  // The body and its section view hold the same paragraph: edit both.
  const document: Document = structuredClone(snapshot.document);
  const body = document.package.document;
  for (const blocks of [
    body.content,
    ...(body.sections ?? []).map((section) => section.content),
  ])
    for (const block of blocks)
      if (block.type === "paragraph" && block.paraId === "10000000")
        block.content = structuredClone(content);
  const saved = await serializeCanonicalSave({
    snapshot: {
      ...snapshot,
      document,
      changedBlockIds: ["10000000"],
      version: snapshot.version + 1,
    },
    featureFlags: { selectiveSave: true },
  });
  const xml =
    (await (
      await ZipArchive.read(saved.buffer)
    )
      .entry("word/document.xml")
      ?.text()) ?? "";
  return {
    text: [...xml.matchAll(/<w:(?:t|delText)\b[^>]*>([^<]*)</gu)]
      .map(([, text]) => text)
      .join(""),
    revisions: [...xml.matchAll(/<w:(ins|del)\b/gu)].map(([, name]) => name),
  };
};

describe("a canonical save with pending suggestions", () => {
  test("writes the document as if every suggestion were rejected", async () => {
    expect(
      await savedBody([
        run("Hello"),
        { type: "insertion", info: info(1, "s1"), content: [run(" there")] },
      ]),
    ).toEqual({ text: "Hello", revisions: [] });
    expect(
      await savedBody([
        run("Hel"),
        { type: "deletion", info: info(1, "s1"), content: [run("lo")] },
      ]),
    ).toEqual({ text: "Hello", revisions: [] });
  });

  test("keeps an ordinary tracked change beside a suggestion", async () => {
    expect(
      await savedBody([
        run("Hello"),
        { type: "insertion", info: info(1, "s1"), content: [run(" there")] },
        { type: "insertion", info: info(2), content: [run("!")] },
      ]),
    ).toEqual({ text: "Hello!", revisions: ["ins"] });
  });
});
