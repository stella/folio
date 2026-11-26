import { describe, expect, spyOn, test } from "bun:test";

import { ZipArchive } from "@stll/docx-core/zip";

import { prepareCanonicalDocxSource } from "./canonicalSessionInput";
import { ensureParaIds } from "./ensureParaIds";
import { ensureUniqueControlIds } from "./ensureUniqueControlIds";
import { ensureUniqueRevisionIds } from "./ensureUniqueRevisionIds";
import { createEmptyDocx } from "./rezip";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const AUTH = 'w:author="Reviewer" w:date="2026-01-01T00:00:00Z"';

const inserted = (id: number, text: string) =>
  `<w:p><w:ins w:id="${String(id)}" ${AUTH}><w:r><w:t>${text}</w:t></w:r></w:ins></w:p>`;
const control = (id: number, text: string) =>
  `<w:sdt><w:sdtPr><w:id w:val="${String(id)}"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:sdtContent></w:sdt>`;

/** A package whose paragraphs name no ids and whose controls and revisions repeat theirs. */
const unnormalizedPackage = async () => {
  const zip = await ZipArchive.read(await createEmptyDocx());
  zip.set(
    "word/document.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="${W}"><w:body>${inserted(5, "a")}${inserted(5, "b")}${control(7, "c")}${control(7, "d")}<w:p><w:r><w:t>e</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`,
  );
  return zip.generate({ compression: "store" });
};

const xmlParts = async (docx: Uint8Array) => {
  const zip = await ZipArchive.read(docx);
  const parts = new Map<string, string>();
  for (const entry of zip.entries())
    if (entry.path.endsWith(".xml") || entry.path.endsWith(".rels"))
      // oxlint-disable-next-line no-await-in-loop -- a handful of small parts
      parts.set(entry.path, await entry.text());
  return parts;
};

describe("canonical package normalization", () => {
  test("writes what the paragraph, control and revision passes write in turn", async () => {
    const source = await unnormalizedPackage();
    const paragraphs = await ensureParaIds(source);
    const chained = await ensureUniqueRevisionIds(
      (await ensureUniqueControlIds(paragraphs.docx)).docx,
    );
    const prepared = await prepareCanonicalDocxSource(source);
    expect(prepared.isOk()).toBe(true);
    if (!prepared.isOk()) return;
    expect(await xmlParts(prepared.value.docx)).toEqual(
      await xmlParts(chained.docx),
    );
    expect([...prepared.value.mintedParaIds]).toEqual([
      ...paragraphs.mintedParaIds,
    ]);
    const body = (await xmlParts(prepared.value.docx)).get("word/document.xml");
    expect(body?.match(/<w:ins w:id="5"/gu)?.length).toBe(1);
    expect(body?.match(/<w:id w:val="7"\/>/gu)?.length).toBe(1);
  });

  test("reads the package once and writes it once, however many passes rewrite it", async () => {
    const source = await unnormalizedPackage();
    const reads = spyOn(ZipArchive, "read");
    const writes = spyOn(ZipArchive.prototype, "generate");
    try {
      const prepared = await prepareCanonicalDocxSource(source);
      expect(prepared.isOk()).toBe(true);
      // One bounded read plans every pass; one more opens the package to write it.
      expect(reads).toHaveBeenCalledTimes(2);
      expect(writes).toHaveBeenCalledTimes(1);
    } finally {
      reads.mockRestore();
      writes.mockRestore();
    }
  });

  test("a package with nothing to normalize is read once and not written", async () => {
    const prepared = await prepareCanonicalDocxSource(
      await unnormalizedPackage(),
    );
    if (!prepared.isOk()) throw prepared.error;
    const reads = spyOn(ZipArchive, "read");
    const writes = spyOn(ZipArchive.prototype, "generate");
    try {
      const again = await prepareCanonicalDocxSource(prepared.value.docx);
      if (!again.isOk()) throw again.error;
      expect(again.value.docx).toEqual(prepared.value.docx);
      expect(reads).toHaveBeenCalledTimes(1);
      expect(writes).not.toHaveBeenCalled();
    } finally {
      reads.mockRestore();
      writes.mockRestore();
    }
  });
});
