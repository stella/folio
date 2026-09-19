/**
 * `w:tblLook` says the same six things twice, and the two spellings differ.
 *
 * `w:val` is a bitmask and the only spelling a producer older than ECMA-376's
 * second edition writes; the per-flag attributes came later and win where both
 * are present. Between them sits a third state the old model could not hold: an
 * explicit `w:lastRow="0"`, which states the region off rather than leaving it
 * to `w:val`'s bit.
 *
 * These cases pin all three ends: the reader keeps what was written, the
 * resolver decides what it means, and the serializer writes it back.
 */

import { describe, expect, test } from "bun:test";
import { resolveTableLook } from "./tableLook";

import type { Document } from "../types/document";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import { serializeTableFormatting } from "./serializer/tableSerializer";
import { parseTableProperties } from "./tableParser";
import { parseXmlDocument, type XmlElement } from "./xmlParser";

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

const lookOf = (attributes: string) => {
  const root = parseXmlDocument(
    `<w:tblPr ${NS}><w:tblLook ${attributes}/></w:tblPr>`,
  ) as XmlElement;
  return parseTableProperties(root)?.look;
};

/** Rebuild the element from the model: verbatim replay would prove nothing. */
const rewritten = (attributes: string): string => {
  const look = lookOf(attributes);
  return serializeTableFormatting(look === undefined ? {} : { look });
};

describe("reading w:tblLook", () => {
  test("keeps w:val as written and invents no flag from its bits", () => {
    expect(lookOf('w:val="04A0"')).toEqual({ val: "04A0" });
  });

  test("keeps an explicit off, which is not the same as saying nothing", () => {
    expect(lookOf('w:val="04A0" w:firstRow="1" w:lastRow="0"')).toEqual({
      val: "04A0",
      firstRow: true,
      lastRow: false,
    });
  });

  test("reads every ST_OnOff spelling of both polarities", () => {
    expect(lookOf('w:firstRow="true" w:lastRow="off" w:noHBand="on" w:noVBand="false"')).toEqual({
      firstRow: true,
      lastRow: false,
      noHBand: true,
      noVBand: false,
    });
  });
});

describe("resolving w:tblLook", () => {
  test("a val-only look resolves from its bits", () => {
    expect(resolveTableLook(lookOf('w:val="04A0"'))).toEqual({
      firstRow: true,
      lastRow: false,
      firstColumn: true,
      lastColumn: false,
      noHBand: false,
      noVBand: true,
    });
  });

  test("a stated flag overrides the bit that disagrees with it", () => {
    // 0x04A0 sets firstRow; the attribute says otherwise and is the later
    // spelling of the same fact.
    expect(resolveTableLook(lookOf('w:val="04A0" w:firstRow="0"')).firstRow).toBe(false);
    expect(resolveTableLook(lookOf('w:val="0000" w:noHBand="1"')).noHBand).toBe(true);
  });

  test("a val outside ST_ShortHexNumber states nothing", () => {
    expect(resolveTableLook({ val: "not-hex" })).toEqual({
      firstRow: false,
      lastRow: false,
      firstColumn: false,
      lastColumn: false,
      noHBand: false,
      noVBand: false,
    });
  });
});

describe("writing w:tblLook", () => {
  test("writes w:val first, then every flag the author stated", () => {
    expect(rewritten('w:val="04A0" w:firstRow="1" w:lastRow="0" w:noVBand="1"')).toContain(
      '<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:noVBand="1"/>',
    );
  });

  test("normalises the spellings of a stated flag without changing its polarity", () => {
    expect(rewritten('w:firstRow="on" w:lastRow="false"')).toContain(
      '<w:tblLook w:firstRow="1" w:lastRow="0"/>',
    );
  });
});

/**
 * A table whose look is a bare `w:val`, under a style that formats the first
 * row and bands the rest. Reading the flags alone would leave every region
 * unstyled, which is what the resolver exists to prevent.
 */
const valOnlyDocument = (look: string): Document => ({
  package: {
    styles: {
      styles: [
        {
          styleId: "Banded",
          type: "table",
          tblStylePr: [
            { type: "firstRow", rPr: { bold: true } },
            { type: "band1Horz", rPr: { italic: true } },
          ],
        },
      ],
    },
    document: {
      content: [
        {
          type: "table",
          formatting: { styleId: "Banded", look: { val: look } },
          rows: [0, 1].map((index) => ({
            cells: [
              {
                content: [
                  {
                    type: "paragraph",
                    content: [
                      { type: "run", content: [{ type: "text", text: `row ${String(index)}` }] },
                    ],
                  },
                ],
              },
            ],
          })),
        },
      ],
    },
  },
});

const rowMarkNames = (document: Document, rowIndex: number): string[] => {
  const table = toProseDoc(document, { styles: document.package.styles }).firstChild;
  const run = table?.child(rowIndex).firstChild?.firstChild?.firstChild;
  return (run?.marks ?? []).map((mark) => mark.type.name);
};

describe("a table that states its look only as w:val", () => {
  test("takes its first-row formatting and its banding from the bits", () => {
    // 0x04A0: firstRow, firstColumn, noVBand. Horizontal banding stays on.
    const document = valOnlyDocument("04A0");
    expect(rowMarkNames(document, 0)).toContain("bold");
    expect(rowMarkNames(document, 1)).toContain("italic");
  });

  test("and turns banding off when the bit says so", () => {
    // 0x0200 is noHBand, so no row takes a band style.
    const document = valOnlyDocument("0200");
    expect(rowMarkNames(document, 0)).not.toContain("italic");
    expect(rowMarkNames(document, 1)).not.toContain("italic");
  });
});
