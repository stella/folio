import JSZip from "jszip";
import { readFile } from "node:fs/promises";
import { beforeAll, describe, expect, test } from "bun:test";

import {
  DocxProjectionError,
  initializeDocxProjection,
  projectCompressedDocx,
  projectCompressedDocxWithReviewFacts,
} from "./projection";
import type { DocxReviewFactsWire } from "./projection";

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

    expect(projection[0]).toBe(3);
    expect(projection[1].map(([, text]) => text)).toEqual(["Before", "Inside"]);
    expect(projection[1][1]?.[4]).toEqual(["table", "table-0", 0, 0]);
  });

  test("preserves every direct text style across the WebAssembly wire", async () => {
    const archive = new JSZip();
    archive.file(
      "word/document.xml",
      `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:rPr><w:b/><w:highlight w:val="yellow"/><w:vertAlign w:val="superscript"/></w:rPr><w:t>Styled</w:t></w:r></w:p></w:body></w:document>`,
    );

    const projection = await projectCompressedDocx(
      await archive.generateAsync({ compression: "DEFLATE", type: "uint8array" }),
    );

    expect(projection[1][0]?.[3]).toEqual([
      [0, 6, "bold"],
      [0, 6, "highlight"],
      [0, 6, "superscript"],
    ]);
  });

  test("exposes direct style identifiers separately from resolved outline levels", async () => {
    const archive = new JSZip();
    archive.file(
      "word/document.xml",
      `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:pPr><w:pStyle w:val="Derived"/></w:pPr><w:r><w:t>Inherited</w:t></w:r></w:p><w:p><w:pPr><w:pStyle w:val="Derived"/><w:outlineLvl w:val="3"/></w:pPr><w:r><w:t>Direct</w:t></w:r></w:p></w:body></w:document>`,
    );
    archive.file(
      "word/styles.xml",
      `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:styleId="Base"><w:pPr><w:outlineLvl w:val="1"/></w:pPr></w:style><w:style w:type="paragraph" w:styleId="Derived"><w:basedOn w:val="Base"/></w:style></w:styles>`,
    );

    const projection = await projectCompressedDocx(
      await archive.generateAsync({ compression: "DEFLATE", type: "uint8array" }),
    );

    expect(projection[1].map((paragraph) => paragraph[5])).toEqual(["Derived", "Derived"]);
    expect(projection[2][4]).toEqual([
      "known",
      [
        [0, 1],
        [1, 3],
      ],
    ]);
  });

  test("wraps malformed packages in a typed boundary error", async () => {
    const projection = projectCompressedDocx(new Uint8Array([1, 2, 3]));

    await expect(projection).rejects.toBeInstanceOf(DocxProjectionError);
  });

  test("returns document and review facts from one package projection", async () => {
    const archive = new JSZip();
    archive.file(
      "word/document.xml",
      `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:ins w:id="7" w:author="Ada"><w:r><w:t>new</w:t></w:r></w:ins></w:p></w:body></w:document>`,
    );
    archive.file(
      "word/comments.xml",
      `<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:comment w:id="1" w:author="Lin"><w:p w14:paraId="AAAAAAAA"/></w:comment></w:comments>`,
    );
    archive.file(
      "word/commentsExtended.xml",
      `<w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"><w15:commentEx w15:paraId="AAAAAAAA" w15:done="1"/></w15:commentsEx>`,
    );
    archive.file(
      "word/_rels/document.xml.rels",
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="comments" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/><Relationship Id="commentsExtended" Type="http://schemas.microsoft.com/office/2011/relationships/commentsExtended" Target="commentsExtended.xml"/></Relationships>`,
    );

    const projection = await projectCompressedDocxWithReviewFacts(
      await archive.generateAsync({ compression: "DEFLATE", type: "uint8array" }),
    );

    expect(projection[0]).toBe(1);
    expect(projection[1][1][0]?.[1]).toBe("new");
    const expectedReviewFacts = [
      1,
      [
        "known",
        [
          [
            "insertion",
            "Ada",
            null,
            "7",
            [
              "known",
              [
                [
                  [0, 0, 0],
                  [0, 3, 3],
                ],
                "new",
                "text",
              ],
            ],
          ],
        ],
      ],
      ["known", [["1", "Lin", null, null, null, "resolved", ["unknown", "unsupported-location"]]]],
    ] as const satisfies DocxReviewFactsWire;
    expect(projection[2]).toEqual(expectedReviewFacts);
  });

  test("materializes footnote markers consistently in paragraphs and comment references", async () => {
    const archive = new JSZip();
    archive.file(
      "word/document.xml",
      `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:commentRangeStart w:id="1"/><w:r><w:t>Title</w:t><w:footnoteReference w:id="1"/></w:r><w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r></w:p></w:body></w:document>`,
    );
    archive.file(
      "word/comments.xml",
      `<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:comment w:id="1" w:author="Ada"><w:p><w:r><w:t>Review</w:t></w:r></w:p></w:comment></w:comments>`,
    );
    archive.file(
      "word/_rels/document.xml.rels",
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="comments" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/></Relationships>`,
    );
    const bytes = await archive.generateAsync({ compression: "DEFLATE", type: "uint8array" });

    const host = await projectCompressedDocxWithReviewFacts(bytes);
    const readable = await projectCompressedDocxWithReviewFacts(bytes, {
      textMaterialization: "readable-plain-text",
    });

    expect(host[1][1][0]?.[1]).toBe("Title\u0002");
    expect(readable[1][1][0]?.[1]).toBe("Title");
    expect(host[2][2]).toEqual([
      "known",
      [
        [
          "1",
          "Ada",
          null,
          null,
          null,
          "open",
          [
            "known",
            [
              [
                [0, 0, 0],
                [0, 6, 6],
              ],
              "Review",
              "Title\u0002",
            ],
          ],
        ],
      ],
    ]);
    expect(readable[2][2]).toEqual([
      "known",
      [
        [
          "1",
          "Ada",
          null,
          null,
          null,
          "open",
          [
            "known",
            [
              [
                [0, 0, 0],
                [0, 5, 5],
              ],
              "Review",
              "Title",
            ],
          ],
        ],
      ],
    ]);
  });
});
