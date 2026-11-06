/**
 * Formatting a container's last paragraph and deleting it in one batch: direct
 * and tracked refuse the same operation, whatever order the batch names them in
 * and whatever else it deletes, and nothing reported applied goes missing.
 */

import { describe, expect, test } from "bun:test";

import { ZipArchive } from "@stll/docx-core/zip";

import { createDocx } from "../docx/rezip";
import { createEmptyDocument } from "../utils/createDocument";
import { FolioDocxReviewer } from "./headless";

const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"';

const ID = (index: number) => `1000000${String(index)}`;

const docx = async (count: number) => {
  const zip = await ZipArchive.read(await createDocx(createEmptyDocument()));
  zip.set(
    "word/document.xml",
    `<?xml version="1.0" encoding="UTF-8"?><w:document ${NS}><w:body>${Array.from(
      { length: count },
      (_, index) =>
        `<w:p w14:paraId="${ID(index + 1)}"><w:r><w:t>T${String(index + 1)}</w:t></w:r></w:p>`,
    ).join("")}</w:body></w:document>`,
  );
  return (await zip.generate()).slice().buffer;
};

const format = (index: number) => ({
  type: "setBlockParagraphProperties",
  blockId: ID(index),
  properties: { alignment: "right" },
});
const remove = (index: number) => ({ type: "deleteBlock", blockId: ID(index) });

const outcome = (reviewer: FolioDocxReviewer) =>
  reviewer
    .getContent()
    .map(({ text, directAlignment }) => `${text}:${directAlignment ?? "-"}`);

const BATCHES: [string, number, Record<string, unknown>[]][] = [
  [
    "format last, then delete all",
    3,
    [format(3), remove(1), remove(2), remove(3)],
  ],
  [
    "delete all, then format last",
    3,
    [remove(1), remove(2), remove(3), format(3)],
  ],
  [
    "format last, delete it first",
    3,
    [format(3), remove(3), remove(1), remove(2)],
  ],
  ["format last, delete both of two", 2, [format(2), remove(1), remove(2)]],
  ["format the only one, delete it", 1, [format(1), remove(1)]],
  [
    "format a middle one, delete all",
    3,
    [format(2), remove(1), remove(2), remove(3)],
  ],
  ["format last, delete last", 3, [format(3), remove(3)]],
];

describe.each(BATCHES)("%s", (_name, count, operations) => {
  const run = async (mode: "direct" | "tracked-changes") => {
    const reviewer = await FolioDocxReviewer.fromBuffer(await docx(count));
    const result = reviewer.applyDocumentOperations(
      {
        version: 1,
        mode,
        operations: operations.map((operation, index) => ({
          id: `op${String(index)}`,
          ...operation,
        })),
      } as never,
      { undefinedReferences: "refuse" },
    );
    return {
      reviewer,
      applied: result.applied.map(({ id }) => id).toSorted(),
      refused: (result.skipped ?? []).map(({ id }) => id).toSorted(),
    };
  };

  test("direct and tracked apply and refuse the same", async () => {
    const direct = await run("direct");
    const tracked = await run("tracked-changes");
    expect(tracked.applied).toEqual(direct.applied);
    expect(tracked.refused).toEqual(direct.refused);
  });

  test("accepting gives the direct result; rejecting, the source", async () => {
    const direct = await run("direct");
    const tracked = await run("tracked-changes");
    const saved = await tracked.reviewer.toBuffer();

    const accepted = await FolioDocxReviewer.fromBuffer(saved);
    accepted.acceptAll();
    expect(outcome(accepted)).toEqual(outcome(direct.reviewer));

    const rejected = await FolioDocxReviewer.fromBuffer(saved);
    rejected.rejectAll();
    expect(outcome(rejected)).toEqual(
      Array.from({ length: count }, (_, index) => `T${String(index + 1)}:-`),
    );
  });
});
