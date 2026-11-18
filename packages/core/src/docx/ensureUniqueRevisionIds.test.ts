import { describe, expect, test } from "bun:test";

import { ZipArchive } from "@stll/docx-core/zip";

import { FolioDocxReviewer } from "../ai-edits/headless";
import { fromMarkdown } from "../markdown/fromMarkdown";
import {
  holdsUniqueSeedIds,
  prepareCanonicalDocumentInput,
  prepareCanonicalDocxInput,
} from "./canonicalSessionInput";
import { ensureUniqueRevisionIds } from "./ensureUniqueRevisionIds";
import { createEmptyDocx } from "./rezip";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const MC = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const AUTH = 'w:author="Reviewer" w:date="2026-01-01T00:00:00Z"';

const inserted = (id: number, text: string) =>
  `<w:ins w:id="${String(id)}" ${AUTH}><w:r><w:t>${text}</w:t></w:r></w:ins>`;
const deleted = (id: number, text: string) =>
  `<w:del w:id="${String(id)}" ${AUTH}><w:r><w:delText>${text}</w:delText></w:r></w:del>`;
const paragraph = (paraId: string, content: string, mark?: string) =>
  `<w:p xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" w14:paraId="${paraId}">${mark === undefined ? "" : `<w:pPr><w:rPr>${mark}</w:rPr></w:pPr>`}${content}</w:p>`;

const packageOf = async (body: string, header?: string) => {
  const zip = await ZipArchive.read(await createEmptyDocx());
  zip.set(
    "word/document.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="${W}" xmlns:mc="${MC}"><w:body>${body}<w:sectPr/></w:body></w:document>`,
  );
  if (header !== undefined)
    zip.set(
      "word/header1.xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr xmlns:w="${W}">${header}</w:hdr>`,
    );
  return zip.generate({ compression: "store" });
};

const partOf = async (docx: Uint8Array, path = "word/document.xml") =>
  (await (await ZipArchive.read(docx)).entry(path)?.text()) ?? "";
const revisionIds = async (docx: Uint8Array, path = "word/document.xml") =>
  [
    ...(await partOf(docx, path)).matchAll(
      /<w:(?:ins|del)\b[^>]*?w:id="(-?\d+)"/gu,
    ),
  ].map(([, id]) => Number(id));

describe("unique revision ids", () => {
  test("the first record keeps its id and later ones get unused ids", async () => {
    const docx = await packageOf(
      paragraph("10000001", `${inserted(5, "a")}${inserted(5, "b")}`) +
        paragraph("10000002", deleted(5, "c"), `<w:ins w:id="1" ${AUTH}/>`),
    );
    const { docx: normalized, reminted } = await ensureUniqueRevisionIds(docx);
    expect(reminted).toBe(2);
    expect(await revisionIds(normalized)).toEqual([5, 2, 1, 3]);
    const xml = await partOf(normalized);
    for (const text of ["<w:t>a</w:t>", "<w:t>b</w:t>", "<w:delText>c"])
      expect(xml).toContain(text);
  });

  test("a record repeated in another story part gets its own id", async () => {
    const docx = await packageOf(
      paragraph("10000001", inserted(5, "a")),
      paragraph("20000001", inserted(5, "h")),
    );
    const { docx: normalized } = await ensureUniqueRevisionIds(docx);
    expect(await revisionIds(normalized)).toEqual([5]);
    expect(await revisionIds(normalized, "word/header1.xml")).toEqual([1]);
  });

  test("range markers sharing a record's id are left alone", async () => {
    const docx = await packageOf(
      paragraph(
        "10000001",
        `<w:bookmarkStart w:id="5" w:name="b"/>${inserted(5, "a")}<w:bookmarkEnd w:id="5"/><w:commentRangeStart w:id="5"/><w:r><w:t>x</w:t></w:r><w:commentRangeEnd w:id="5"/>`,
      ),
    );
    const result = await ensureUniqueRevisionIds(docx);
    expect(result.reminted).toBe(0);
    expect(result.docx).toBe(docx);
  });

  test("alternate content restating a record keeps its id", async () => {
    const docx = await packageOf(
      paragraph(
        "10000001",
        `<mc:AlternateContent><mc:Choice Requires="w14">${inserted(5, "a")}</mc:Choice><mc:Fallback>${inserted(5, "a")}</mc:Fallback></mc:AlternateContent>`,
      ),
    );
    const result = await ensureUniqueRevisionIds(docx);
    expect(result.reminted).toBe(0);
    expect(result.docx).toBe(docx);
  });

  test("normalizing twice changes nothing more", async () => {
    const docx = await packageOf(
      paragraph("10000001", `${inserted(5, "a")}${inserted(5, "b")}`),
    );
    const once = await ensureUniqueRevisionIds(docx);
    const twice = await ensureUniqueRevisionIds(once.docx);
    expect(twice.reminted).toBe(0);
    expect(twice.docx).toBe(once.docx);
  });
});

// A table whose inserted rows, cell words and cell marks all repeat one id
// each, as some producers write them; every record resolves as its own change.
const repeatedTable = () => {
  const cell = (text: string) =>
    `<w:tc><w:p xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:pPr><w:rPr><w:ins w:id="90" ${AUTH}/></w:rPr></w:pPr>${inserted(80, text)}</w:p></w:tc>`;
  const row = (a: string, b: string) =>
    `<w:tr><w:trPr><w:ins w:id="70" ${AUTH}/></w:trPr>${cell(a)}${cell(b)}</w:tr>`;
  return `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol/><w:gridCol/></w:tblGrid>${row("A1", "B1")}${row("A2", "B2")}</w:tbl>`;
};

test("a canonical reviewer opens a package repeating revision ids, resolves every record and saves it stable", async () => {
  const docx = await packageOf(
    paragraph("10000001", "<w:r><w:t>Intro.</w:t></w:r>") +
      paragraph(
        "10000002",
        "<w:r><w:t>Alpha.</w:t></w:r>",
        `<w:del w:id="91" ${AUTH}/>`,
      ) +
      repeatedTable() +
      paragraph("10000003", "<w:r><w:t>Final.</w:t></w:r>"),
  );
  const prepared = (await prepareCanonicalDocxInput(docx)).unwrap();
  const accepting = await FolioDocxReviewer.fromBuffer(
    prepared.slice().buffer,
    {
      session: "canonical",
    },
  );
  expect(accepting.acceptAll().skipped).toEqual([]);
  expect(accepting.getContent().map(({ text }) => text)).toEqual([
    "Intro.",
    "Alpha.A1",
    "B1",
    "A2",
    "B2",
    "Final.",
  ]);
  const rejecting = await FolioDocxReviewer.fromBuffer(
    prepared.slice().buffer,
    {
      session: "canonical",
    },
  );
  expect(rejecting.rejectAll().skipped).toEqual([]);
  expect(rejecting.getContent().map(({ text }) => text)).toEqual([
    "Intro.",
    "Alpha.",
    "Final.",
  ]);
  // The saved package already holds unique ids: preparing it again is a no-op.
  const pending = await FolioDocxReviewer.fromBuffer(prepared.slice().buffer, {
    session: "canonical",
  });
  const saved = new Uint8Array(await pending.toBuffer());
  expect((await ensureUniqueRevisionIds(saved)).reminted).toBe(0);
  const reopened = await FolioDocxReviewer.fromBuffer(saved.slice().buffer, {
    session: "canonical",
  });
  expect(reopened.getChanges().map(({ id }) => id)).toEqual(
    pending.getChanges().map(({ id }) => id),
  );
});

test("a model whose records share a revision id seeds a canonical session with ids of their own", async () => {
  const document = fromMarkdown("Alpha.\n\nBravo.");
  const info = { id: 7, author: "Reviewer", date: "2026-01-01T00:00:00Z" };
  for (const block of document.package.document.content)
    if (block.type === "paragraph")
      block.content = [{ type: "insertion", info, content: block.content }];
  expect(holdsUniqueSeedIds(document)).toBe(false);
  const prepared = (await prepareCanonicalDocumentInput(document)).unwrap();
  expect(holdsUniqueSeedIds(prepared)).toBe(true);
  const ids = prepared.package.document.content.flatMap((block) =>
    block.type === "paragraph"
      ? block.content.flatMap((item) =>
          item.type === "insertion" ? [item.info.id] : [],
        )
      : [],
  );
  expect(new Set(ids).size).toBe(2);
});
