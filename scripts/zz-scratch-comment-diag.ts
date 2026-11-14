import { ZipArchive } from "@stll/docx-core/zip";
import { parseDocx } from "@stll/folio-core/docx/parser";
import { repackDocx } from "@stll/folio-core/docx/rezip";

import { prepareOpDocument } from "./lib/corpus-invariants/op-sequences";

const file = Bun.argv[2] ?? "";
const bytes = new Uint8Array(await Bun.file(file).arrayBuffer());
const buffer = bytes.slice().buffer;
const zip = await ZipArchive.read(bytes);
console.log("DIAG parts", zip.paths().join(","));
const comments = (await zip.entry("word/comments.xml")?.text()) ?? "";
console.log(
  "DIAG comments.xml",
  comments.replaceAll(/\s+/gu, " ").slice(0, 3000),
);
const main = (await zip.entry("word/document.xml")?.text()) ?? "";
for (const match of main.matchAll(/<w:comment[A-Za-z]*\b[^>]*>/gu)) {
  const at = match.index ?? 0;
  console.log(
    "DIAG marker",
    match[0],
    "CONTEXT",
    main.slice(Math.max(0, at - 700), at + 300).replaceAll(/\s+/gu, " "),
  );
}
const parsed = await parseDocx(buffer, { preloadFonts: false });
const ids = (d: typeof parsed) =>
  JSON.stringify((d.package.document.comments ?? []).map(({ id }) => id));
console.log("DIAG parsed comments", ids(parsed));
try {
  await repackDocx(parsed, { updateModifiedDate: false });
  console.log("DIAG parsed repack ok");
} catch (error) {
  console.log(
    "DIAG parsed repack FAILS",
    String(error).replaceAll(/\s+/gu, " ").slice(0, 1500),
  );
}
const document = await prepareOpDocument({
  bytes,
  buffer,
  parsed,
  documentPart: "word/document.xml",
  budgetMs: 600_000,
});
console.log("DIAG prepared comments", ids(document));
try {
  await repackDocx(document, { updateModifiedDate: false });
  console.log("DIAG prepared repack ok");
} catch (error) {
  console.log(
    "DIAG prepared repack FAILS",
    String(error).replaceAll(/\s+/gu, " ").slice(0, 1500),
  );
}
