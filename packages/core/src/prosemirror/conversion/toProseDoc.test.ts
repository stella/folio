import { describe, expect, test } from "bun:test";

import type { Document, TableCell, Theme } from "../../types/document";
import { toProseDoc } from "./toProseDoc";

const officeTheme: Theme = {
  colorScheme: {
    dk1: "000000",
    lt1: "FFFFFF",
    dk2: "44546A",
    lt2: "E7E6E6",
    accent1: "4472C4",
    accent2: "ED7D31",
    accent3: "A5A5A5",
    accent4: "FFC000",
    accent5: "5B9BD5",
    accent6: "70AD47",
    hlink: "0563C1",
    folHlink: "954F72",
  },
};

function tableCellBorderColor(
  attrs: Record<string, unknown>,
  side: "top" | "bottom" | "left" | "right",
): { rgb?: string; themeColor?: string } | undefined {
  const borders = attrs["borders"] as
    | Record<string, { color?: { rgb?: string; themeColor?: string } }>
    | null
    | undefined;
  return borders?.[side]?.color;
}

function firstTableCellAttrs(doc: Document): Record<string, unknown> {
  const pmDoc = toProseDoc(doc, {
    ...(doc.package.styles ? { styles: doc.package.styles } : {}),
    ...(doc.package.theme !== undefined ? { theme: doc.package.theme } : {}),
  });
  const firstTable = pmDoc.firstChild;
  const firstRow = firstTable?.firstChild;
  const firstCell = firstRow?.firstChild;
  return (firstCell?.attrs ?? {}) as Record<string, unknown>;
}

describe("toProseDoc", () => {
  test("tracks docDefaults spacing so empty paragraphs retain it during layout", () => {
    const document: Document = {
      package: {
        document: {
          content: [{ type: "paragraph", content: [] }],
        },
        styles: {
          docDefaults: { pPr: { spaceAfter: 160 } },
          styles: [],
        },
      },
    };

    const doc = toProseDoc(document, { styles: document.package.styles });

    expect(doc.firstChild?.attrs.spaceAfter).toBe(160);
    expect(doc.firstChild?.attrs.spacingFromDocDefaults).toEqual({ after: true });
    expect(doc.firstChild?.attrs.spacingExplicit).toBeNull();
  });

  test("applies oneNDA paragraph mark defaults to unformatted visible text like Word", () => {
    const document: Document = {
      package: {
        document: {
          content: [
            {
              type: "paragraph",
              formatting: {
                runProperties: {
                  bold: true,
                  fontSize: 21,
                  fontFamily: { ascii: "Arial", hAnsi: "Arial" },
                },
              },
              content: [
                {
                  type: "run",
                  formatting: {},
                  content: [{ type: "text", text: "TERMS" }],
                },
              ],
            },
          ],
        },
      },
    };

    const doc = toProseDoc(document);
    const text = doc.firstChild?.firstChild;

    expect(text?.marks.some((mark) => mark.type.name === "bold")).toBe(true);
    expect(text?.marks.some((mark) => mark.type.name === "italic")).toBe(false);
    expect(text?.marks.find((mark) => mark.type.name === "fontSize")?.attrs.size).toBe(21);
    expect(text?.marks.find((mark) => mark.type.name === "fontFamily")?.attrs.ascii).toBe("Arial");
  });

  test("keeps oneNDA heading direct run formatting ahead of paragraph mark defaults", () => {
    const document: Document = {
      package: {
        document: {
          content: [
            {
              type: "paragraph",
              formatting: {
                runProperties: {
                  italic: true,
                  fontSize: 18,
                },
              },
              content: [
                {
                  type: "run",
                  formatting: {
                    bold: true,
                    fontSize: 20,
                  },
                  content: [{ type: "text", text: "PARTIES AND " }],
                },
              ],
            },
          ],
        },
      },
    };

    const doc = toProseDoc(document);
    const text = doc.firstChild?.firstChild;

    expect(text?.marks.some((mark) => mark.type.name === "bold")).toBe(true);
    expect(text?.marks.some((mark) => mark.type.name === "italic")).toBe(false);
    expect(text?.marks.find((mark) => mark.type.name === "fontSize")?.attrs.size).toBe(20);
  });

  test("keeps paragraph font size when a run only overrides the font family", () => {
    const document: Document = {
      package: {
        document: {
          content: [
            {
              type: "paragraph",
              formatting: {
                runProperties: {
                  fontSize: 16,
                  fontFamily: { ascii: "Arial Narrow", hAnsi: "Arial Narrow" },
                },
              },
              content: [
                {
                  type: "run",
                  formatting: {
                    fontFamily: { ascii: "Arial Narrow", hAnsi: "Arial Narrow" },
                  },
                  content: [{ type: "text", text: " " }],
                },
              ],
            },
          ],
        },
      },
    };

    const doc = toProseDoc(document);
    const text = doc.firstChild?.firstChild;

    expect(text?.marks.find((mark) => mark.type.name === "fontSize")?.attrs.size).toBe(16);
  });

  test("keeps the default paragraph style size when paragraph and run only set a font family", () => {
    const document: Document = {
      package: {
        document: {
          content: [
            {
              type: "paragraph",
              formatting: {
                runProperties: {
                  fontFamily: { ascii: "Arial", hAnsi: "Arial" },
                },
              },
              content: [
                {
                  type: "run",
                  formatting: {
                    fontFamily: { ascii: "Arial", hAnsi: "Arial" },
                  },
                  content: [{ type: "text", text: "Default style text" }],
                },
              ],
            },
          ],
        },
        styles: {
          docDefaults: { rPr: { fontSize: 22 } },
          styles: [
            {
              styleId: "Normal",
              type: "paragraph",
              default: true,
              rPr: { fontSize: 24 },
            },
          ],
        },
      },
    };

    const doc = toProseDoc(document, { styles: document.package.styles });
    const paragraph = doc.firstChild;
    const text = paragraph?.firstChild;

    expect(paragraph?.attrs.defaultTextFormatting.fontSize).toBe(24);
    expect(text?.marks.find((mark) => mark.type.name === "fontSize")?.attrs.size).toBe(24);
    expect(text?.marks.find((mark) => mark.type.name === "fontFamily")?.attrs.ascii).toBe("Arial");
  });

  test("does not leak a paragraph-mark-only size into directly formatted body text", () => {
    const document: Document = {
      package: {
        document: {
          content: [
            {
              type: "paragraph",
              formatting: {
                runProperties: {
                  fontSize: 52,
                  fontSizeCs: 52,
                  fontFamily: { ascii: "Arial", hAnsi: "Arial" },
                },
              },
              content: [
                {
                  type: "run",
                  formatting: {
                    fontFamily: { ascii: "Arial", hAnsi: "Arial" },
                  },
                  content: [{ type: "text", text: "Body text" }],
                },
              ],
            },
          ],
        },
        styles: {
          docDefaults: { rPr: { fontSize: 22, fontSizeCs: 22 } },
          styles: [
            {
              styleId: "Normal",
              type: "paragraph",
              default: true,
              rPr: { fontSize: 24, fontSizeCs: 24 },
            },
          ],
        },
      },
    };

    const doc = toProseDoc(document, { styles: document.package.styles });
    const paragraph = doc.firstChild;
    const text = paragraph?.firstChild;

    expect(paragraph?.attrs.defaultTextFormatting.fontSize).toBe(52);
    expect(text?.marks.find((mark) => mark.type.name === "fontSize")?.attrs.size).toBe(24);
  });

  test("keeps named paragraph style fonts ahead of paragraph-mark formatting", () => {
    const document: Document = {
      package: {
        document: {
          content: [
            {
              type: "paragraph",
              formatting: {
                styleId: "Header",
                runProperties: {
                  fontSize: 11,
                  fontFamily: { ascii: "Calibri", hAnsi: "Calibri" },
                },
              },
              content: [
                {
                  type: "run",
                  formatting: {},
                  content: [{ type: "text", text: "   " }],
                },
                {
                  type: "run",
                  formatting: {
                    fontSize: 11,
                    fontFamily: { ascii: "Calibri", hAnsi: "Calibri" },
                  },
                  content: [{ type: "text", text: "reference" }],
                },
              ],
            },
          ],
        },
        styles: {
          styles: [
            {
              styleId: "Normal",
              type: "paragraph",
              default: true,
              rPr: {
                fontSize: 12,
                fontFamily: { ascii: "Times New Roman", hAnsi: "Times New Roman" },
              },
            },
            {
              styleId: "Header",
              type: "paragraph",
              basedOn: "Normal",
              rPr: {
                fontSize: 12,
                fontFamily: { ascii: "Times New Roman", hAnsi: "Times New Roman" },
              },
            },
          ],
        },
      },
    };

    const doc = toProseDoc(document, { styles: document.package.styles });
    const spaces = doc.firstChild?.firstChild;
    const reference = doc.firstChild?.maybeChild(1);

    expect(spaces?.marks.find((mark) => mark.type.name === "fontSize")?.attrs.size).toBe(12);
    expect(spaces?.marks.find((mark) => mark.type.name === "fontFamily")?.attrs.ascii).toBe(
      "Times New Roman",
    );
    expect(reference?.marks.find((mark) => mark.type.name === "fontSize")?.attrs.size).toBe(11);
    expect(reference?.marks.find((mark) => mark.type.name === "fontFamily")?.attrs.ascii).toBe(
      "Calibri",
    );
  });

  test("does not inherit paragraph mark all-caps onto visible text", () => {
    const document: Document = {
      package: {
        document: {
          content: [
            {
              type: "paragraph",
              formatting: {
                runProperties: {
                  allCaps: true,
                },
              },
              content: [
                {
                  type: "run",
                  formatting: {
                    italic: true,
                  },
                  content: [{ type: "text", text: "mandated lead arranger" }],
                },
              ],
            },
          ],
        },
      },
    };

    const doc = toProseDoc(document);
    const text = doc.firstChild?.firstChild;

    expect(text?.marks.some((mark) => mark.type.name === "allCaps")).toBe(false);
    expect(text?.marks.some((mark) => mark.type.name === "italic")).toBe(true);
  });

  test("preserves explicit run-level all-caps", () => {
    const document: Document = {
      package: {
        document: {
          content: [
            {
              type: "paragraph",
              content: [
                {
                  type: "run",
                  formatting: {
                    allCaps: true,
                    italic: true,
                  },
                  content: [{ type: "text", text: "mandated lead arranger" }],
                },
              ],
            },
          ],
        },
      },
    };

    const doc = toProseDoc(document);
    const text = doc.firstChild?.firstChild;

    expect(text?.marks.some((mark) => mark.type.name === "allCaps")).toBe(true);
    expect(text?.marks.some((mark) => mark.type.name === "italic")).toBe(true);
  });

  test("preserves direct run marks on tab nodes", () => {
    const document: Document = {
      package: {
        document: {
          content: [
            {
              type: "paragraph",
              content: [
                {
                  type: "run",
                  formatting: {
                    underline: { style: "single" },
                  },
                  content: [{ type: "tab" }],
                },
              ],
            },
          ],
        },
      },
    };

    const doc = toProseDoc(document);
    const tab = doc.firstChild?.firstChild;

    expect(tab?.type.name).toBe("tab");
    expect(tab?.marks.find((mark) => mark.type.name === "underline")?.attrs.style).toBe("single");
  });

  test("does not leak paragraph mark underline onto directly formatted text runs", () => {
    const document: Document = {
      package: {
        document: {
          content: [
            {
              type: "paragraph",
              formatting: {
                runProperties: {
                  underline: { style: "single" },
                },
              },
              content: [
                {
                  type: "run",
                  formatting: { fontSize: 22 },
                  content: [{ type: "text", text: "By:" }],
                },
                {
                  type: "run",
                  formatting: {
                    underline: { style: "single" },
                    fontSize: 22,
                  },
                  content: [{ type: "tab" }],
                },
              ],
            },
          ],
        },
      },
    };

    const doc = toProseDoc(document);
    const label = doc.firstChild?.child(0);
    const tab = doc.firstChild?.child(1);

    expect(label?.marks.some((mark) => mark.type.name === "underline")).toBe(false);
    expect(
      label?.marks.find((mark) => mark.type.name === "runFormattingOverride")?.attrs.underline,
    ).toBe("none");
    expect(tab?.marks.find((mark) => mark.type.name === "underline")?.attrs.style).toBe("single");
  });

  test("does not leak paragraph mark bold onto directly formatted highlighted text", () => {
    const document: Document = {
      package: {
        document: {
          content: [
            {
              type: "paragraph",
              formatting: {
                runProperties: {
                  bold: true,
                },
              },
              content: [
                {
                  type: "run",
                  formatting: {
                    fontSize: 22,
                    highlight: "yellow",
                  },
                  content: [{ type: "text", text: "COMPANY NAME" }],
                },
              ],
            },
          ],
        },
      },
    };

    const doc = toProseDoc(document);
    const text = doc.firstChild?.firstChild;

    expect(text?.marks.some((mark) => mark.type.name === "bold")).toBe(false);
    expect(text?.marks.find((mark) => mark.type.name === "runFormattingOverride")?.attrs.bold).toBe(
      false,
    );
  });

  test("does not leak paragraph mark character spacing onto a directly formatted run", () => {
    const document: Document = {
      package: {
        document: {
          content: [
            {
              type: "paragraph",
              formatting: {
                runProperties: { spacing: -60 },
              },
              content: [
                {
                  type: "run",
                  formatting: { boldCs: true },
                  content: [{ type: "text", text: "Signature name" }],
                },
              ],
            },
          ],
        },
      },
    };

    const doc = toProseDoc(document);
    const paragraph = doc.firstChild;
    const text = paragraph?.firstChild;

    expect(paragraph?.attrs.defaultTextFormatting.spacing).toBe(-60);
    expect(text?.marks.find((mark) => mark.type.name === "characterSpacing")?.attrs.spacing).toBe(
      0,
    );
  });

  test("applies oneNDA table style run properties before direct run properties", () => {
    const document: Document = {
      package: {
        styles: {
          styles: [
            {
              styleId: "TableStyle",
              type: "table",
              rPr: {
                bold: true,
                fontFamily: {
                  ascii: "Open Sans Light",
                  hAnsi: "Open Sans Light",
                },
              },
            },
          ],
        },
        document: {
          content: [
            {
              type: "table",
              formatting: { styleId: "TableStyle" },
              rows: [
                {
                  cells: [
                    {
                      content: [
                        {
                          type: "paragraph",
                          content: [
                            {
                              type: "run",
                              content: [{ type: "text", text: "Styled cell" }],
                            },
                          ],
                        },
                      ],
                    },
                  ],
                },
              ],
            },
          ],
        },
      },
    };

    const doc = toProseDoc(document, { styles: document.package.styles });
    const text = doc.firstChild?.firstChild?.firstChild?.firstChild?.firstChild;

    expect(text?.marks.some((mark) => mark.type.name === "bold")).toBe(true);
    expect(text?.marks.find((mark) => mark.type.name === "fontFamily")?.attrs.ascii).toBe(
      "Open Sans Light",
    );
  });

  test("resolves themed table-cell border colors against the document theme", () => {
    const cellFormatting: TableCell["formatting"] = {
      borders: {
        top: {
          style: "single",
          size: 8,
          color: { themeColor: "accent2" },
        },
      },
    };
    const document: Document = {
      package: {
        theme: officeTheme,
        document: {
          content: [
            {
              type: "table",
              rows: [
                {
                  cells: [
                    {
                      formatting: cellFormatting,
                      content: [{ type: "paragraph", content: [] }],
                    },
                  ],
                },
              ],
            },
          ],
        },
      },
    };

    const attrs = firstTableCellAttrs(document);

    expect(tableCellBorderColor(attrs, "top")?.rgb).toBe("ED7D31");
    expect(tableCellBorderColor(attrs, "top")?.themeColor).toBeUndefined();
  });

  test("accepts unknown table-cell border styles preserved from OOXML", () => {
    const document: Document = {
      package: {
        document: {
          content: [
            {
              type: "table",
              rows: [
                {
                  cells: [
                    {
                      formatting: {
                        borders: {
                          bottom: {
                            style: "0",
                            size: 0,
                            color: { rgb: "000000" },
                          },
                        },
                      },
                      content: [{ type: "paragraph", content: [] }],
                    },
                  ],
                },
              ],
            },
          ],
        },
      },
    };

    const attrs = firstTableCellAttrs(document);
    const borders = attrs["borders"] as { bottom?: { style?: string } } | undefined;

    expect(borders?.bottom?.style).toBe("0");
  });

  test("resolves themed table-cell border color with themeTint to the modified RGB", () => {
    const document: Document = {
      package: {
        theme: officeTheme,
        document: {
          content: [
            {
              type: "table",
              rows: [
                {
                  type: "tableRow",
                  cells: [
                    {
                      type: "tableCell",
                      formatting: {
                        borders: {
                          left: {
                            style: "single",
                            size: 4,
                            color: {
                              themeColor: "accent1",
                              themeTint: "33",
                            },
                          },
                        },
                      },
                      content: [{ type: "paragraph", content: [] }],
                    },
                  ],
                },
              ],
            },
          ],
        },
      },
    };

    const attrs = firstTableCellAttrs(document);

    // accent1 (4472C4) blended toward white with tint 0x33/0xFF.
    expect(tableCellBorderColor(attrs, "left")?.rgb).toBe("DAE3F3");
    expect(tableCellBorderColor(attrs, "left")?.themeColor).toBeUndefined();
  });

  test("passes plain RGB and auto table-cell border colors through unchanged", () => {
    const document: Document = {
      package: {
        theme: officeTheme,
        document: {
          content: [
            {
              type: "table",
              rows: [
                {
                  type: "tableRow",
                  cells: [
                    {
                      type: "tableCell",
                      formatting: {
                        borders: {
                          top: {
                            style: "single",
                            size: 8,
                            color: { rgb: "FF0000" },
                          },
                          bottom: {
                            style: "single",
                            size: 8,
                            color: { rgb: "auto" },
                          },
                        },
                      },
                      content: [{ type: "paragraph", content: [] }],
                    },
                  ],
                },
              ],
            },
          ],
        },
      },
    };

    const attrs = firstTableCellAttrs(document);

    expect(tableCellBorderColor(attrs, "top")?.rgb).toBe("FF0000");
    expect(tableCellBorderColor(attrs, "bottom")?.rgb).toBe("auto");
  });

  test("keeps unresolved themed table-cell border colors when no document theme exists", () => {
    const document: Document = {
      package: {
        document: {
          content: [
            {
              type: "table",
              rows: [
                {
                  cells: [
                    {
                      formatting: {
                        borders: {
                          left: {
                            style: "single",
                            size: 8,
                            color: { themeColor: "accent1", themeTint: "33" },
                          },
                        },
                      },
                      content: [{ type: "paragraph", content: [] }],
                    },
                  ],
                },
              ],
            },
          ],
        },
      },
    };

    const attrs = firstTableCellAttrs(document);

    expect(tableCellBorderColor(attrs, "left")?.themeColor).toBe("accent1");
  });

  test("applies built-in Word Normal defaults when styles.xml is absent", () => {
    const document: Document = {
      package: {
        document: {
          content: [
            {
              type: "paragraph",
              content: [
                {
                  type: "run",
                  content: [{ type: "text", text: "Plain paragraph" }],
                },
              ],
            },
          ],
        },
      },
    };

    const doc = toProseDoc(document);
    const paragraph = doc.firstChild;

    expect(paragraph?.attrs.spaceAfter).toBe(160);
    expect(paragraph?.attrs.lineSpacing).toBe(259);
    expect(paragraph?.attrs.lineSpacingRule).toBe("auto");
  });

  test("keeps paragraph style run defaults above default character style", () => {
    const document: Document = {
      package: {
        styles: {
          styles: [
            {
              styleId: "Normal",
              type: "paragraph",
              default: true,
              rPr: {
                fontFamily: { ascii: "Arial", hAnsi: "Arial" },
              },
            },
            {
              styleId: "DefaultChar",
              type: "character",
              default: true,
              rPr: {
                fontFamily: { ascii: "Cambria", hAnsi: "Cambria" },
              },
            },
          ],
        },
        document: {
          content: [
            {
              type: "paragraph",
              formatting: { styleId: "Normal" },
              content: [
                {
                  type: "run",
                  content: [{ type: "text", text: "Plain paragraph" }],
                },
              ],
            },
          ],
        },
      },
    };

    const doc = toProseDoc(document, { styles: document.package.styles });
    const defaultTextFormatting = doc.firstChild?.attrs.defaultTextFormatting;

    expect(defaultTextFormatting?.fontFamily?.ascii).toBe("Arial");
    expect(defaultTextFormatting?.fontFamily?.hAnsi).toBe("Arial");
  });

  test("uses default character style when paragraph style has no run defaults", () => {
    const document: Document = {
      package: {
        styles: {
          styles: [
            {
              styleId: "Normal",
              type: "paragraph",
              default: true,
            },
            {
              styleId: "DefaultChar",
              type: "character",
              default: true,
              rPr: {
                fontFamily: { ascii: "Cambria", hAnsi: "Cambria" },
              },
            },
          ],
        },
        document: {
          content: [
            {
              type: "paragraph",
              formatting: { styleId: "Normal" },
              content: [
                {
                  type: "run",
                  content: [{ type: "text", text: "Plain paragraph" }],
                },
              ],
            },
          ],
        },
      },
    };

    const doc = toProseDoc(document, { styles: document.package.styles });
    const defaultTextFormatting = doc.firstChild?.attrs.defaultTextFormatting;

    expect(defaultTextFormatting?.fontFamily?.ascii).toBe("Cambria");
    expect(defaultTextFormatting?.fontFamily?.hAnsi).toBe("Cambria");
  });

  test("preserves DOCX field instruction and cached display text", () => {
    const document: Document = {
      package: {
        document: {
          content: [
            {
              type: "paragraph",
              content: [
                {
                  type: "simpleField",
                  instruction: ' MERGEFIELD "Client Name" \\* MERGEFORMAT ',
                  fieldType: "MERGEFIELD",
                  content: [
                    {
                      type: "run",
                      content: [{ type: "text", text: "Acme s.r.o." }],
                    },
                  ],
                },
              ],
            },
          ],
        },
      },
    };

    const doc = toProseDoc(document);
    const field = doc.firstChild?.firstChild;

    expect(field?.type.name).toBe("field");
    expect(field?.attrs.fieldType).toBe("MERGEFIELD");
    expect(field?.attrs.instruction).toBe(' MERGEFIELD "Client Name" \\* MERGEFORMAT ');
    expect(field?.attrs.displayText).toBe("Acme s.r.o.");
    expect(field?.attrs.fieldKind).toBe("simple");
  });

  test("does not apply TOC paragraph-mark font defaults to field result text", () => {
    const document: Document = {
      package: {
        styles: {
          styles: [
            {
              styleId: "Normal",
              type: "paragraph",
              default: true,
              rPr: {
                fontFamily: { ascii: "CG Times", hAnsi: "CG Times" },
              },
            },
            {
              styleId: "TOC1",
              type: "paragraph",
              basedOn: "Normal",
              rPr: {
                fontFamily: { ascii: "CG Times", hAnsi: "CG Times" },
              },
            },
          ],
        },
        document: {
          content: [
            {
              type: "paragraph",
              formatting: {
                styleId: "TOC1",
                runProperties: {
                  fontFamily: { ascii: "Calibri", hAnsi: "Calibri" },
                },
              },
              content: [
                {
                  type: "complexField",
                  fieldType: "TOC",
                  instruction: ' TOC \\o "1-1" ',
                  fieldResult: [
                    {
                      type: "run",
                      formatting: {},
                      content: [{ type: "text", text: "Definitions and Interpretation" }],
                    },
                  ],
                },
              ],
            },
          ],
        },
      },
    };

    const doc = toProseDoc(document, { styles: document.package.styles });
    const paragraph = doc.firstChild;
    const field = paragraph?.firstChild;
    const fieldFont = field?.marks.find((mark) => mark.type.name === "fontFamily");
    const defaultTextFormatting = paragraph?.attrs.defaultTextFormatting;

    expect(field?.type.name).toBe("field");
    expect(fieldFont?.attrs.ascii).toBe("CG Times");
    expect(defaultTextFormatting?.fontFamily?.ascii).toBe("CG Times");
  });

  test("converts inline content controls without synthetic marks", () => {
    const document: Document = {
      package: {
        document: {
          content: [
            {
              type: "paragraph",
              content: [
                {
                  type: "inlineSdt",
                  properties: { sdtType: "plainText" },
                  content: [
                    {
                      type: "run",
                      content: [{ type: "text", text: "Controlled" }],
                    },
                  ],
                },
              ],
            },
          ],
        },
      },
    };

    const doc = toProseDoc(document);
    const sdt = doc.firstChild?.firstChild;

    expect(sdt?.type.name).toBe("sdt");
    expect(sdt?.marks).toEqual([]);
    expect(sdt?.firstChild?.text).toBe("Controlled");
  });

  test("anchors point comments to nearby text for display", () => {
    const document: Document = {
      package: {
        document: {
          content: [
            {
              type: "paragraph",
              content: [
                {
                  type: "run",
                  content: [{ type: "text", text: "Commented text" }],
                },
                { type: "commentReference", id: 42 },
              ],
            },
          ],
        },
      },
    };

    const doc = toProseDoc(document);
    const text = doc.firstChild?.firstChild;
    const commentMark = text?.marks.find((mark) => mark.type.name === "comment");

    expect(commentMark?.attrs.commentId).toBe(42);
  });

  test("applies active comment ranges to every text-emitting inline branch", () => {
    const document: Document = {
      package: {
        document: {
          content: [
            {
              type: "paragraph",
              content: [
                { type: "commentRangeStart", id: 99 },
                {
                  type: "run",
                  content: [{ type: "text", text: "plain" }],
                },
                {
                  type: "hyperlink",
                  href: "https://stella.law",
                  children: [
                    {
                      type: "run",
                      content: [{ type: "text", text: "link" }],
                    },
                  ],
                },
                {
                  type: "simpleField",
                  instruction: " PAGE ",
                  fieldType: "PAGE",
                  content: [
                    {
                      type: "run",
                      content: [{ type: "text", text: "1" }],
                    },
                  ],
                },
                {
                  type: "insertion",
                  info: { id: 1, author: "User" },
                  content: [
                    {
                      type: "run",
                      content: [{ type: "text", text: "inserted" }],
                    },
                  ],
                },
                {
                  type: "mathEquation",
                  display: "inline",
                  ommlXml: "<m:oMath />",
                  plainText: "x",
                },
                { type: "commentRangeEnd", id: 99 },
              ],
            },
          ],
        },
      },
    };

    const doc = toProseDoc(document);
    const paragraph = doc.firstChild;
    const commentMarkedNodeTypes: string[] = [];

    for (let index = 0; index < (paragraph?.childCount ?? 0); index++) {
      const node = paragraph?.child(index);
      if (!node) {
        continue;
      }
      if (node.marks.some((mark) => mark.type.name === "comment" && mark.attrs.commentId === 99)) {
        commentMarkedNodeTypes.push(node.type.name);
      }
    }

    expect(commentMarkedNodeTypes).toEqual(["text", "text", "text"]);
  });
});
