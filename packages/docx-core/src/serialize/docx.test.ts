import { describe, expect, test } from "bun:test";
import JSZip from "jszip";

import type { Document, Table } from "../model/document";
import { serializeDocumentToDocx } from "./docx";

const docWithBorder = (style: string, rgb: string): Document => {
  const table: Table = {
    type: "table",
    rows: [
      {
        type: "tableRow",
        cells: [
          {
            type: "tableCell",
            content: [
              {
                type: "paragraph",
                content: [
                  {
                    type: "run",
                    content: [{ type: "text", text: "x" }],
                  },
                ],
              },
            ],
          },
        ],
      },
    ],
    formatting: {
      borders: { top: { style, size: 4, color: { rgb } } },
    },
  };

  return {
    package: {
      document: {
        content: [table],
      },
    },
  };
};

const readDocumentXml = async (buf: ArrayBuffer): Promise<string> => {
  const zip = await JSZip.loadAsync(buf);
  const xml = await zip.file("word/document.xml")?.async("string");
  if (xml === undefined) {
    throw new Error("word/document.xml missing from serialized docx");
  }
  return xml;
};

const readNumberingXml = async (buf: ArrayBuffer): Promise<string> => {
  const zip = await JSZip.loadAsync(buf);
  const xml = await zip.file("word/numbering.xml")?.async("string");
  if (xml === undefined) {
    throw new Error("word/numbering.xml missing from serialized docx");
  }
  return xml;
};

test("numbering markers serialize independent complex-script typography", async () => {
  const document: Document = {
    package: {
      document: { content: [] },
      numbering: {
        abstractNums: [
          {
            abstractNumId: 1,
            levels: [
              {
                ilvl: 0,
                numFmt: "arabicAlpha",
                lvlText: "%1.",
                rPr: {
                  fontFamily: { ascii: "Arial", cs: "Traditional Arabic" },
                  fontSize: 20,
                  fontSizeCs: 32,
                  bold: true,
                  boldCs: false,
                  italic: false,
                  italicCs: true,
                  cs: false,
                },
              },
            ],
          },
        ],
        nums: [{ numId: 1, abstractNumId: 1 }],
      },
    },
  };

  const xml = await readNumberingXml(await serializeDocumentToDocx(document));
  expect(xml).toContain('w:cs="Traditional Arabic"');
  expect(xml).toContain('<w:bCs w:val="0"/>');
  expect(xml).toContain('<w:i w:val="0"/>');
  expect(xml).toContain("<w:iCs/>");
  expect(xml).toContain('<w:szCs w:val="32"/>');
  expect(xml).toContain('<w:cs w:val="0"/>');
});

describe("DOCX border serialization escapes attribute values", () => {
  test("a style value cannot inject extra XML attributes", async () => {
    const xml = await readDocumentXml(
      await serializeDocumentToDocx(docWithBorder('single" w:evil="1', "CCCCCC")),
    );
    // The injected attribute must not appear as a real attribute, and the
    // quote must be entity-escaped.
    expect(xml).not.toContain('w:evil="1"');
    expect(xml).toContain("&quot;");
  });

  test("a color value cannot break out of its attribute", async () => {
    const xml = await readDocumentXml(
      await serializeDocumentToDocx(docWithBorder("single", 'FF0000" x="y')),
    );
    expect(xml).not.toContain('x="y"');
    expect(xml).toContain("&quot;");
  });

  test("ordinary border values serialize unescaped", async () => {
    const xml = await readDocumentXml(
      await serializeDocumentToDocx(docWithBorder("single", "CCCCCC")),
    );
    expect(xml).toContain('w:val="single"');
    expect(xml).toContain('w:color="CCCCCC"');
  });
});

describe("DOCX positional tab serialization", () => {
  test("writes positional tab attributes", async () => {
    const document: Document = {
      package: {
        document: {
          content: [
            {
              type: "paragraph",
              content: [
                {
                  type: "run",
                  content: [
                    {
                      type: "tab",
                      positional: {
                        relativeTo: "margin",
                        alignment: "right",
                        leader: "dot",
                      },
                    },
                  ],
                },
              ],
            },
          ],
        },
      },
    };

    const xml = await readDocumentXml(await serializeDocumentToDocx(document));
    expect(xml).toContain('<w:ptab w:relativeTo="margin" w:alignment="right" w:leader="dot"/>');
  });
});
