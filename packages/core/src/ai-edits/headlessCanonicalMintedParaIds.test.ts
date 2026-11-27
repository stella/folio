/**
 * A paragraph id the package did not carry is minted when a canonical
 * reviewer opens it. Until a save writes it, the id names a position, not
 * the paragraph: every read reports it positional, through edits and their
 * resolution. The save writes it, so the reopened package carries it as its
 * own and reads it stable. An id the package carried reads stable throughout.
 */

import { describe, expect, test } from "bun:test";

import { ZipArchive } from "@stll/docx-core/zip";

import { FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION } from "../document-operations";
import { prepareCanonicalDocxSource } from "../docx/canonicalSessionInput";
import { createDocx } from "../docx/rezip";
import type { Paragraph } from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";
import { FolioDocxReviewer } from "./headless";

const AUTHORED_PARA_ID = "1A2B3C4D";

const paragraph = (text: string, paraId?: string): Paragraph => ({
  type: "paragraph",
  content: [{ type: "run", content: [{ type: "text", text }] }],
  ...(paraId !== undefined && { paraId }),
});

/**
 * A package whose first paragraph names no id and whose second names one. A
 * save gives every paragraph an id, so the saved ids but the authored one are
 * taken off again.
 */
const mixedPackage = async (): Promise<ArrayBuffer> => {
  const template = createEmptyDocument();
  const saved = await createDocx({
    ...template,
    package: {
      ...template.package,
      document: {
        ...template.package.document,
        content: [
          paragraph("Minted source text."),
          paragraph("Authored paragraph.", AUTHORED_PARA_ID),
        ],
      },
    },
  });
  const zip = await ZipArchive.read(saved);
  const documentXml = await zip.entry("word/document.xml")?.text();
  if (documentXml === undefined) throw new Error("the package has no body");
  zip.set(
    "word/document.xml",
    documentXml.replace(
      new RegExp(
        `\\s(?:w14:paraId|w14:textId)="(?!${AUTHORED_PARA_ID})[^"]*"`,
        "gu",
      ),
      "",
    ),
  );
  return (await zip.generate({ compression: "store" })).buffer as ArrayBuffer;
};

const open = (buffer: ArrayBuffer): Promise<FolioDocxReviewer> =>
  FolioDocxReviewer.fromBuffer(buffer.slice(0), {
    author: "Canonical Reviewer",
    session: "canonical",
  });

/** Each block's id and stability, as the content and snapshot readers report them. */
const stabilities = (reviewer: FolioDocxReviewer) => {
  const content = reviewer
    .getContent()
    .map(({ id, idStability }) => [id, idStability ?? "stable"]);
  expect(
    reviewer
      .snapshot()
      .blocks.map(({ id, idStability }) => [id, idStability ?? "stable"]),
  ).toEqual(content);
  return content;
};

describe("a paragraph id a canonical reviewer mints on opening", () => {
  test("reads positional until a save writes it, then stable", async () => {
    const source = await mixedPackage();
    const prepared = await prepareCanonicalDocxSource(source);
    if (prepared.isErr()) throw prepared.error;
    const [minted] = prepared.value.mintedParaIds;
    expect(prepared.value.mintedParaIds.size).toBe(1);
    if (minted === undefined) throw new Error("Expected a minted id");

    const reviewer = await open(source);
    expect(stabilities(reviewer)).toEqual([
      [minted, "positional"],
      [AUTHORED_PARA_ID, "stable"],
    ]);

    const receipt = reviewer.applyDocumentOperations(
      {
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        mode: "tracked-changes",
        operations: [
          {
            id: "replace-minted",
            type: "replaceInBlock",
            blockId: minted,
            find: "source",
            replace: "edited",
          },
          {
            id: "replace-authored",
            type: "replaceInBlock",
            blockId: AUTHORED_PARA_ID,
            find: "Authored",
            replace: "Named",
          },
        ],
      },
      { undefinedReferences: "refuse" },
    );
    expect(receipt.skipped).toEqual([]);
    reviewer.acceptAll();
    expect(reviewer.getContent().map(({ text }) => text)).toEqual([
      "Minted edited text.",
      "Named paragraph.",
    ]);
    expect(stabilities(reviewer)).toEqual([
      [minted, "positional"],
      [AUTHORED_PARA_ID, "stable"],
    ]);

    const reopened = await open(await reviewer.toBuffer());
    expect(stabilities(reopened)).toEqual([
      [minted, "stable"],
      [AUTHORED_PARA_ID, "stable"],
    ]);
  });
});
