import { expect, test } from "bun:test";

import { ZipArchive } from "@stll/docx-core/zip";

import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
} from "../document-operations";
import { ensureParaIds } from "../docx/ensureParaIds";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";
import { FolioDocxReviewer } from "./headless";

type Mode = "direct" | "tracked-changes";
const STYLED = [
  "# Agreement",
  "## Scope",
  "### Definitions",
  "The Supplier delivers the goods.",
  "The Buyer pays on delivery.",
  "## Payment",
  "Payment is due in ten days (see clause 1).",
  "Closing paragraph.",
].join("\n\n");
const PLAIN = [
  "# Service Agreement",
  "This agreement is made between the parties named below.",
  "The Supplier delivers the goods on time and in good order.",
  "The Buyer pays each invoice within thirty days.",
  "Signed in two copies.",
].join("\n\n");
const LISTS = [
  "# Delivery Terms",
  "The following apply to every order.",
  "- Goods are packed securely\n- Goods are insured in transit",
  "Payment happens in stages.",
  "1. Deposit on signature\n2. Balance on delivery\n3. Retention after inspection",
  "Closing remarks.",
].join("\n\n");

const open = async (bytes: Uint8Array | ArrayBuffer) =>
  FolioDocxReviewer.fromBuffer(bytes, { author: "T" });
const fixture = async (md: string) =>
  (await ensureParaIds(await createDocx(fromMarkdown(md)))).docx;
const body = async (bytes: ArrayBuffer) => {
  const xml =
    (await (
      await ZipArchive.read(new Uint8Array(bytes))
    )
      .entry("word/document.xml")
      ?.text()) ?? "";
  return (xml.split("<w:body>")[1] ?? "")
    .replace(/ (w14|w15|w16\w*):\w+="[^"]*"/g, "")
    .replace(/ w:rsid\w*="[^"]*"/g, "")
    .replace(/ w:(author|date)="[^"]*"/g, "")
    .replace(/<w:sectPr[\s\S]*$/, "")
    .replace(/<w:p[ >]/g, "\n$&");
};
const rows = (r: FolioDocxReviewer) =>
  r
    .snapshot()
    .blocks.map((b) => `${JSON.stringify(b.text)} ${b.styleId ?? ""}`);

const scenario = async (
  name: string,
  md: string,
  steps: (id: string[]) => FolioDocumentOperation[][],
) => {
  const bytes = await fixture(md);
  const out: Record<string, unknown> = {};
  for (const mode of ["direct", "tracked-changes"] as Mode[]) {
    const r = await open(bytes);
    const id = r.snapshot().blocks.map((b) => b.id);
    for (const ops of steps(id)) {
      const res = r.applyDocumentOperations(
        {
          version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
          mode,
          atomic: false,
          operations: ops,
        },
        { undefinedReferences: "refuse" },
      );
      out[`${mode} skipped`] = [
        ...((out[`${mode} skipped`] as unknown[]) ?? []),
        ...res.skipped,
      ];
    }
    const saved = await r.toBuffer();
    if (mode === "tracked-changes") out["tracked xml"] = await body(saved);
    const live = await open(saved);
    out[`${mode} pending`] = rows(live);
    live.acceptAll();
    const acceptedBytes = await live.toBuffer();
    if (mode === "tracked-changes")
      out["accepted xml"] = await body(acceptedBytes);
    out[`${mode} accepted`] = rows(await open(acceptedBytes));
    out[`${mode} comments`] = (await open(acceptedBytes))
      .getComments()
      .map((c) => c.anchoredText);
  }
  console.log(
    `=== ${name}\n${Object.entries(out)
      .map(
        ([k, v]) =>
          `-- ${k}:\n${typeof v === "string" ? v : JSON.stringify(v, null, 1)}`,
      )
      .join("\n")}`,
  );
  expect(out["tracked-changes accepted"]).toEqual(out["direct accepted"]);
};

const comment = (blockId: string): FolioDocumentOperation => ({
  id: "c",
  type: "commentOnBlock",
  blockId,
  comment: { text: "A note." },
});
const del = (blockId: string): FolioDocumentOperation => ({
  id: "d",
  type: "deleteBlock",
  blockId,
});

test("debug A", async () => {
  await scenario("A", STYLED, (id) => [
    [del(id[7]!)],
    [
      {
        id: "a",
        type: "insertAfterBlock",
        blockId: id[6]!,
        text: "Term goods party written.",
        styleId: "Heading2",
      },
      {
        id: "b",
        type: "insertBeforeBlock",
        blockId: id[6]!,
        text: "Term schedule term.",
      },
    ],
  ]);
});
test("debug B styled", async () => {
  await scenario("B styled", STYLED, (id) => [
    [
      {
        id: "r",
        type: "replaceInBlock",
        blockId: id[3]!,
        find: "goods",
        replace: "revised",
      },
    ],
    [comment(id[3]!)],
    [del(id[3]!)],
    [del(id[4]!)],
  ]);
});
test("debug B plain", async () => {
  await scenario("B plain", PLAIN, (id) => [
    [comment(id[1]!)],
    [del(id[1]!)],
    [del(id[2]!)],
  ]);
});
test("debug B plain full", async () => {
  await scenario("B plain full", PLAIN, (id) => [
    [del(id[0]!)],
    [comment(id[1]!)],
    [del(id[1]!)],
    [del(id[2]!)],
  ]);
});
test("debug B lists", async () => {
  await scenario("B lists", LISTS, (id) => [
    [comment(id[2]!)],
    [del(id[2]!)],
    [del(id[3]!)],
  ]);
});
test("debug B plain no comment", async () => {
  await scenario("B plain no comment", PLAIN, (id) => [
    [del(id[1]!)],
    [del(id[2]!)],
  ]);
});
