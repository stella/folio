import JSZip from "jszip";
import { readFile } from "node:fs/promises";
import { beforeAll, describe, expect, test } from "bun:test";

import { DocxProjectionError, initializeDocxProjection, projectCompressedDocx } from "./projection";

const documentXml = `<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>Before</w:t></w:r></w:p>
    <w:tbl><w:tr><w:tc><w:p><w:r><w:t>Inside</w:t></w:r></w:p></w:tc></w:tr></w:tbl>
  </w:body>
</w:document>`;

const createDocument = async (): Promise<Uint8Array> => {
  const archive = new JSZip();
  archive.file("word/document.xml", documentXml);
  return archive.generateAsync({ compression: "DEFLATE", type: "uint8array" });
};

beforeAll(async () => {
  const wasm = await readFile(new URL("./generated/docx_kernel_bg.wasm", import.meta.url));
  await initializeDocxProjection({ wasm });
});

describe("DOCX projection TypeScript binding", () => {
  test("runs the versioned Rust projection through WebAssembly", async () => {
    const projection = await projectCompressedDocx(await createDocument());

    expect(projection[0]).toBe(2);
    expect(projection[1].map(([, text]) => text)).toEqual(["Before", "Inside"]);
    expect(projection[1][1]?.[4]).toEqual(["table", "table-0", 0, 0]);
  });

  test("wraps malformed packages in a typed boundary error", async () => {
    const projection = projectCompressedDocx(new Uint8Array([1, 2, 3]));

    await expect(projection).rejects.toBeInstanceOf(DocxProjectionError);
  });
});
