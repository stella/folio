import { expect, spyOn, test } from "bun:test";

import { FOLIO_XML_RESOURCE_LIMITS } from "@stll/docx-core/xml/limits";

import * as documentParser from "./documentParser";
import * as paragraphs from "./paragraphParser";
import { buildStructuralDocumentPatch } from "./structuralXmlPatch";
import * as xml from "./xmlParser";

const wrap = (content: string, background = "") =>
  `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:id="http://schemas.microsoft.com/office/word/2010/wordml">${background}<w:body>${content}</w:body></w:document>`;
const section = (top: number) =>
  `<w:sectPr><w:pgMar w:top="${top}"/></w:sectPr>`;
const paragraph = (id: number, content: string) =>
  `<w:p id:paraId="${id.toString(16).padStart(8, "0")}">${content}</w:p>`;
const text = (value: string) => `<w:r><w:t>${value}</w:t></w:r>`;
const tracked = (id: number) =>
  paragraph(
    id,
    `<w:pPr><w:rPr><w:ins w:id="${id * 2}" w:author="Agent"/></w:rPr></w:pPr><w:ins w:id="${id * 2 + 1}" w:author="Agent">${text("added")}</w:ins>`,
  );

test("plain structural patch compares metadata without modeling each paragraph", () => {
  const prefix = paragraph(1, text("kept"));
  const body = Array.from({ length: 32 }, (_, index) =>
    tracked(index + 2),
  ).join("");
  const parse = spyOn(paragraphs, "parseParagraph");
  const read = spyOn(xml, "parseXmlDocument");
  try {
    const originalXml = wrap(prefix + section(1440));
    const serializedXml = wrap(prefix + body + section(1440));
    const result = buildStructuralDocumentPatch({
      originalXml,
      serializedXml,
      changedIds: new Set(),
    });
    expect(result.type).toBe("patched");
    expect(parse).toHaveBeenCalledTimes(0);
    expect(read).toHaveBeenCalledWith(originalXml);
    expect(read).toHaveBeenCalledWith(serializedXml);
    if (result.type === "patched") {
      expect(read).toHaveBeenCalledWith(result.xml);
      expect(result.xml).toContain(prefix);
      expect(result.xml).toContain(`id:paraId="00000021"`);
    }
  } finally {
    read.mockRestore();
    parse.mockRestore();
  }
});

test("a leading section and following paragraph section still require full parse", () => {
  const first = paragraph(1, text("kept"));
  const second = (top: number) =>
    paragraph(2, `<w:pPr>${section(top)}</w:pPr>${text("section end")}`);
  const parse = spyOn(paragraphs, "parseParagraph");
  try {
    const result = buildStructuralDocumentPatch({
      originalXml: wrap(section(1440) + first + second(1500)),
      serializedXml: wrap(section(1440) + first + second(1600) + tracked(3)),
      changedIds: new Set(),
    });
    expect(result).toEqual({
      type: "declined",
      reason: "structuralBarrierChanged",
    });
    expect(parse.mock.calls.length).toBeGreaterThan(0);
  } finally {
    parse.mockRestore();
  }
});

test("the plain metadata comparison agrees with the full reader through metadata mutations", () => {
  const content = paragraph(1, text("kept"));
  const current = content + tracked(2);
  const background = '<w:background w:color="40A050"/>';
  const changed = '<w:background w:color="1155AA"/>';
  const cases = [
    { source: wrap(content), serialized: wrap(current), result: "patched" },
    ...["leading", "trailing"].flatMap((position) => {
      const attach = (blocks: string, top: number) =>
        position === "leading" ? section(top) + blocks : blocks + section(top);
      const source = wrap(attach(content, 1440), background);
      return [
        {
          source,
          serialized: wrap(attach(current, 1440), background),
          result: "patched",
        },
        {
          source,
          serialized: wrap(attach(current, 1440), background)
            .replaceAll("w:", "doc:")
            .replace("xmlns:w=", "xmlns:doc="),
          result: "patched",
        },
        {
          source,
          serialized: wrap(attach(current, 1600), background),
          result: "declined",
        },
        {
          source,
          serialized: wrap(attach(current, 1440), changed),
          result: "declined",
        },
        { source, serialized: wrap(current, background), result: "declined" },
      ];
    }),
  ];
  for (const { source, serialized, result } of cases) {
    const options = {
      originalXml: source,
      serializedXml: serialized,
      changedIds: new Set<string>(),
    };
    const actual = buildStructuralDocumentPatch(options);
    expect(actual.type).toBe(result);
    const plain = spyOn(
      documentParser,
      "isPlainDocumentMetadataParagraph",
    ).mockReturnValue(false);
    try {
      expect(buildStructuralDocumentPatch(options)).toEqual(actual);
    } finally {
      plain.mockRestore();
    }
  }
});

test("the structural XML reader still declines nesting beyond its budget", () => {
  const { maxDepth } = FOLIO_XML_RESOURCE_LIMITS;
  const nested =
    "<w:r>".repeat(maxDepth) + "<w:t>text</w:t>" + "</w:r>".repeat(maxDepth);
  expect(
    buildStructuralDocumentPatch({
      originalXml: wrap(paragraph(1, text("kept"))),
      serializedXml: wrap(paragraph(1, text("kept")) + paragraph(2, nested)),
      changedIds: new Set(),
    }),
  ).toEqual({ type: "declined", reason: "structuralUnreadableBody" });
});
