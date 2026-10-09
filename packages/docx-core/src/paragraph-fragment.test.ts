import { beforeAll, expect, test } from "bun:test";
import { readFile } from "node:fs/promises";

import {
  DocxProjectionError,
  initializeDocxProjection,
  projectMainDocumentXml,
  projectParagraphFragment,
} from "./projection";
import type { DocxProjectionStructure } from "./projection";

const packageNamespace = "http://schemas.microsoft.com/office/2006/xmlPackage";
const documentNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const relationshipNamespace = "http://schemas.openxmlformats.org/package/2006/relationships";
const mainRelationship =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";

const fragment = (paragraphs: string) =>
  `<flat:package xmlns:flat="${packageNamespace}" xmlns:doc="${documentNamespace}" xmlns:rel="${relationshipNamespace}">
    <flat:part flat:name="/_rels/.rels" flat:contentType="application/xml"><flat:xmlData><rel:Relationships><rel:Relationship Id="main" Type="${mainRelationship}" Target="content/main.xml"/></rel:Relationships></flat:xmlData></flat:part>
    <flat:part flat:name="/content/main.xml" flat:contentType="application/xml"><flat:xmlData><doc:document><doc:body>${paragraphs}</doc:body></doc:document></flat:xmlData></flat:part>
  </flat:package>`;

beforeAll(async () => {
  await initializeDocxProjection({
    wasm: await readFile(new URL("./generated/docx_kernel_bg.wasm", import.meta.url)),
  });
});

test("fragment preserves text and direct formatting with explicit partial evidence", async () => {
  const bytes = new TextEncoder().encode(
    fragment(
      '<doc:p><doc:r><doc:rPr><doc:b/><doc:highlight doc:val="yellow"/></doc:rPr><doc:t>A &amp; 😀</doc:t></doc:r></doc:p>',
    ),
  );
  const partial = await projectParagraphFragment(bytes);
  const full = await projectMainDocumentXml(bytes);
  expect(partial[0]).toBe(full[0]);
  expect(partial[1]).toEqual(full[1]);
  expect(partial[1][0][1]).toBe("A & 😀");
  expect(partial[1][0][3]).toEqual([
    [0, 6, "bold"],
    [0, 6, "highlight"],
  ]);
  expect(partial[2]).toEqual(full[2].map(() => ["unknown", "paragraph-fragment"]));
  expect(partial[3]).toEqual(full[3]);
  expect(partial[4]).toEqual(["incomplete", "styles-part-unavailable"]);
});

test.each(["", "<doc:p/><doc:p/>"])(
  "rejects a package without exactly one paragraph: %s",
  async (paragraphs) => {
    await expect(
      projectParagraphFragment(new TextEncoder().encode(fragment(paragraphs))),
    ).rejects.toBeInstanceOf(DocxProjectionError);
  },
);

test("fragment cannot publish document-relative table coordinates", async () => {
  const bytes = new TextEncoder().encode(
    fragment("<doc:tbl><doc:tr><doc:tc><doc:p/></doc:tc></doc:tr></doc:tbl>"),
  );
  const full = await projectMainDocumentXml(bytes);
  expect(full[1][0]?.[4]).toEqual(["table", "table-0", 0, 0]);
  const partial = await projectParagraphFragment(bytes);
  expect(partial[1][0][4]).toEqual([] as const satisfies DocxProjectionStructure);
  expect(partial[2]).toEqual(full[2].map(() => ["unknown", "paragraph-fragment"]));
});
