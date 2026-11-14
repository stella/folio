import { parseDocx } from "@stll/folio-core/docx/parser";
import { repackDocx } from "@stll/folio-core/docx/rezip";

import {
  generateOpSequence,
  OP_SEQUENCE_SEEDS,
  prepareOpDocument,
  seedFromBytes,
} from "./lib/corpus-invariants/op-sequences";

const file = Bun.argv[2] ?? "";
const bytes = new Uint8Array(await Bun.file(file).arrayBuffer());
const buffer = bytes.slice().buffer;
const parsed = await parseDocx(buffer, { preloadFonts: false });
const input = {
  bytes,
  buffer,
  parsed,
  documentPart: "word/document.xml",
  budgetMs: 600_000,
};
const document = await prepareOpDocument(input);
const ids = (d: typeof document) =>
  JSON.stringify((d.package.document.comments ?? []).map(({ id }) => id));
console.log("DIAG comments", ids(document));
try {
  await repackDocx(document, { updateModifiedDate: false });
  console.log("DIAG prepared repack ok");
} catch (error) {
  console.log("DIAG prepared repack FAILS", String(error).slice(0, 800));
}
const seed = seedFromBytes(bytes);
for (const salt of OP_SEQUENCE_SEEDS) {
  const sequence = generateOpSequence(document, seed ^ salt);
  console.log(
    "DIAG seq",
    salt,
    sequence.steps.map(({ op }) => op.type).join(","),
  );
  for (const [index, step] of sequence.steps.entries()) {
    try {
      // oxlint-disable-next-line no-await-in-loop -- scratch diagnosis
      await repackDocx(step.edit.document, { updateModifiedDate: false });
    } catch (error) {
      console.log(
        "DIAG FAIL",
        salt,
        index,
        step.op.type,
        JSON.stringify(step.op).slice(0, 2000),
      );
      console.log(
        "DIAG before",
        ids(step.before),
        "after",
        ids(step.edit.document),
      );
      console.log(
        "DIAG touched",
        JSON.stringify(step.edit.touched).slice(0, 600),
      );
      console.log("DIAG error", String(error).slice(0, 1200));
      break;
    }
  }
}
