import { expect, test } from "bun:test";

import { ZipArchive } from "@stll/docx-core/zip";

import { createCanonicalSession } from "../canonical/canonicalSession";
import { prepareCanonicalDocxInput } from "../docx/canonicalSessionInput";
import { parseDocx } from "../docx/parser";
import { createEmptyDocx } from "../docx/rezip";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const control = (id: number, text: string) =>
  `<w:sdt><w:sdtPr><w:id w:val="${String(id)}"/></w:sdtPr><w:sdtContent><w:r><w:t>${text}</w:t></w:r></w:sdtContent></w:sdt>`;

test("the canonical session opens a package that repeats a control id, keeping both controls", async () => {
  const zip = await ZipArchive.read(await createEmptyDocx());
  zip.set(
    "word/document.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="${W}"><w:body><w:p>${control(5, "a")}</w:p><w:p>${control(5, "b")}</w:p><w:sectPr/></w:body></w:document>`,
  );
  const docx = await zip.generate({ compression: "store" });
  const session = createCanonicalSession(
    await parseDocx(
      (await prepareCanonicalDocxInput(docx)).unwrap().slice().buffer,
      { preloadFonts: false, detectVariables: false },
    ),
  );
  if (session.isErr()) throw session.error;
  const controls = session.value.document.package.document.content.flatMap(
    (block) =>
      block.type === "paragraph"
        ? block.content.flatMap((item) =>
            item.type === "inlineSdt" ? [item.properties.id] : [],
          )
        : [],
  );
  expect(controls).toHaveLength(2);
  expect(new Set(controls).size).toBe(2);
});
