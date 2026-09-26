import { expect, test } from "bun:test";
import { buildStructuralDocumentPatch } from "./structuralXmlPatch";
import { getChildElements, parseXmlDocument } from "./xmlParser";

const wrap = (body: string, prefix = "w") =>
  `<${prefix}:document xmlns:${prefix}="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:id="http://schemas.microsoft.com/office/word/2010/wordml"><${prefix}:body>${body}</${prefix}:body></${prefix}:document>`;
const paragraph = (id: number, text = `Paragraph ${id}`, prefix = "w") =>
  `<${prefix}:p id:paraId="${id.toString(16).padStart(8, "0")}"><${prefix}:r><${prefix}:t>${text}</${prefix}:t></${prefix}:r></${prefix}:p>`;
const ids = (xml: string) => {
  const root = parseXmlDocument(xml);
  const body = getChildElements(root).at(0);
  return getChildElements(body).map((element) => element.attributes?.["id:paraId"]);
};

test("all deletion subsets and insertion gaps retain surviving source bytes and ordering", () => {
  const original = wrap([1, 2, 3, 4].map((id) => paragraph(id)).join("\n<!-- untouched gap -->\n"));
  for (let mask = 0; mask < 16; mask++) {
    const survivors = [1, 2, 3, 4].filter((id) => (mask & (1 << (id - 1))) !== 0);
    for (let gap = 0; gap <= survivors.length; gap++) {
      const desired = survivors.toSpliced(gap, 0, 5, 6, 7);
      const result = buildStructuralDocumentPatch({
        originalXml: original,
        serializedXml: wrap(desired.map((id) => paragraph(id)).join("")),
        changedIds: new Set(),
      });
      expect(result).not.toBeNull();
      if (result === null) continue;
      expect(ids(result)).toEqual(desired.map((id) => id.toString(16).padStart(8, "0")));
      for (const id of survivors) expect(result).toContain(paragraph(id));
      expect(result.match(/<!-- untouched gap -->/gu)?.length).toBe(3);
    }
  }
});

test("resolves aliased source elements and quoted delimiters without touching source gaps", () => {
  const original = wrap(
    `${paragraph(1, "one", "doc")}<!-- <doc:p> fake -->${paragraph(2, "two", "doc")}`,
    "doc",
  );
  const result = buildStructuralDocumentPatch({
    originalXml: original,
    serializedXml: wrap(paragraph(1, "edited &gt; one") + paragraph(3) + paragraph(2)),
    changedIds: new Set(["00000001"]),
  });
  expect(result).toContain(paragraph(2, "two", "doc"));
  expect(result).toContain("<!-- <doc:p> fake -->");
  expect(result).toContain(
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"',
  );
  expect(result).toContain("edited &gt; one");
});

test("refuses moved, duplicate, missing, or reserved source identities", () => {
  for (const source of [paragraph(1) + paragraph(1), paragraph(0), "<w:p/>"]) {
    expect(
      buildStructuralDocumentPatch({
        originalXml: wrap(source),
        serializedXml: wrap(paragraph(2)),
        changedIds: new Set(),
      }),
    ).toBeNull();
  }
  expect(
    buildStructuralDocumentPatch({
      originalXml: wrap(paragraph(1) + paragraph(2)),
      serializedXml: wrap(paragraph(2) + paragraph(1)),
      changedIds: new Set(),
    }),
  ).toBeNull();
});

test("retains identical tables as barriers but refuses crossing or editing them", () => {
  const table = `<w:tbl><w:tr><w:tc>${paragraph(9)}</w:tc></w:tr></w:tbl>`;
  const original = wrap(paragraph(1) + table + paragraph(2));
  const patch = (body: string) =>
    buildStructuralDocumentPatch({
      originalXml: original,
      serializedXml: wrap(body),
      changedIds: new Set(),
    });
  expect(patch(paragraph(1) + paragraph(3) + table + paragraph(2))).toContain(table);
  expect(patch(paragraph(2) + table + paragraph(1))).toBeNull();
  expect(patch(paragraph(1) + table.replace("Paragraph 9", "edited") + paragraph(2))).toBeNull();
  expect(patch(paragraph(1) + table + paragraph(9))).toBeNull();
});

test("refuses range markers, relationship payloads, section endpoints and Strict splices", () => {
  for (const element of [
    "sectPr",
    "commentRangeStart",
    "bookmarkStart",
    "drawing",
    "footnoteReference",
  ]) {
    const original = wrap(paragraph(1).replace("</w:p>", `<w:${element}/></w:p>`));
    expect(
      buildStructuralDocumentPatch({
        originalXml: original,
        serializedXml: wrap(paragraph(1) + paragraph(2)),
        changedIds: new Set(["00000001"]),
      }),
    ).toBeNull();
  }
  expect(
    buildStructuralDocumentPatch({
      originalXml: wrap(paragraph(1)).replace(
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        "http://purl.oclc.org/ooxml/wordprocessingml/main",
      ),
      serializedXml: wrap(paragraph(1) + paragraph(2)),
      changedIds: new Set(),
    }),
  ).toBeNull();
});

test("untouched image and section paragraphs remain verbatim beside inserted paragraphs", () => {
  for (const element of ["drawing", "sectPr"]) {
    const kept = paragraph(1).replace("</w:p>", `<w:${element}/></w:p>`);
    const result = buildStructuralDocumentPatch({
      originalXml: wrap(kept),
      serializedXml: wrap(kept + paragraph(2)),
      changedIds: new Set(["00000002"]),
    });
    expect(result).toContain(kept);
    expect(result).toContain("Paragraph 2");
  }
});

test("refuses ranges whose endpoints in separate tables enclose a body insertion", () => {
  const table = (id: number, marker: string) =>
    `<w:tbl><w:tr><w:tc>${paragraph(id).replace("</w:p>", marker + "</w:p>")}</w:tc></w:tr></w:tbl>`;
  const start = table(1, '<w:commentRangeStart w:id="1"/>');
  const end = table(2, '<w:commentRangeEnd w:id="1"/>');
  expect(
    buildStructuralDocumentPatch({
      originalXml: wrap(start + end),
      serializedXml: wrap(start + paragraph(3) + end),
      changedIds: new Set(["00000003"]),
    }),
  ).toBeNull();
});
