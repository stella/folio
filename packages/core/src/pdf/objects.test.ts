import { describe, expect, test } from "bun:test";
import {
  createPdfDocument,
  formatNumber,
  pdfArray,
  pdfAsciiString,
  pdfDict,
  pdfName,
  pdfNumber,
  pdfTextString,
} from "./objects";

const latin1 = (bytes: Uint8Array): string => new TextDecoder("latin1").decode(bytes);

const serializeOne = (
  value: Parameters<ReturnType<typeof createPdfDocument>["add"]>[0],
): string => {
  const document = createPdfDocument();
  const root = document.add(value);
  const info = document.add(pdfDict([]));
  return latin1(document.serialize({ rootRef: root, infoRef: info, idHex: () => "00" }));
};

describe("formatNumber", () => {
  test("prints integers without a fraction", () => {
    expect(formatNumber(0)).toBe("0");
    expect(formatNumber(1)).toBe("1");
    expect(formatNumber(100)).toBe("100");
    expect(formatNumber(-42)).toBe("-42");
  });

  test("trims trailing zeros but keeps significant decimals", () => {
    expect(formatNumber(1.5)).toBe("1.5");
    expect(formatNumber(0.75)).toBe("0.75");
    expect(formatNumber(1 / 3)).toBe("0.3333");
  });

  test("never emits a signed zero", () => {
    expect(formatNumber(-0)).toBe("0");
    expect(formatNumber(-1e-9)).toBe("0");
    expect(formatNumber(-0.00004)).toBe("0");
    // A signed zero that arrived by arithmetic, which is how one reaches a
    // coordinate: `Object.is` proves it really is -0 before formatting.
    const computed = [-1].reduce((product, value) => product * value, 0);
    expect(Object.is(computed, -0)).toBe(true);
    expect(formatNumber(computed)).toBe("0");
  });

  test("never emits exponent notation", () => {
    const values = [1e-9, -1e-9, 1e9, -1e9, 0.00001, 1234567.891, -1234567.891, 1e-21];
    for (const value of values) {
      const formatted = formatNumber(value);
      expect(formatted).not.toContain("e");
      expect(formatted).not.toContain("E");
      expect(formatted).not.toBe("-0");
    }
  });

  test("round trips through Number without changing the value it prints", () => {
    for (const value of [12.3456, -7.891, 0.0001, 99999.9999]) {
      expect(Number(formatNumber(value))).toBeCloseTo(value, 4);
    }
  });
});

describe("value serialization", () => {
  test("escapes name characters outside the regular set", () => {
    expect(serializeOne(pdfName("A B#C"))).toContain("/A#20B#23C");
  });

  test("writes a text string as UTF-16BE with a byte order mark", () => {
    expect(serializeOne(pdfTextString("Ω"))).toContain("<FEFF03A9>");
  });

  test("escapes parentheses and backslashes in a byte string", () => {
    expect(serializeOne(pdfAsciiString("a(b)c\\d"))).toContain("(a\\(b\\)c\\\\d)");
  });

  test("drops dictionary entries whose value is undefined", () => {
    const text = serializeOne(
      pdfDict([
        ["Kept", pdfNumber(1)],
        ["Dropped", undefined],
      ]),
    );
    expect(text).toContain("/Kept 1");
    expect(text).not.toContain("/Dropped");
  });

  test("keeps array order", () => {
    expect(serializeOne(pdfArray([pdfNumber(3), pdfNumber(1), pdfNumber(2)]))).toContain("[3 1 2]");
  });
});

describe("file structure", () => {
  test("writes a header, an xref table and a trailer", () => {
    const text = serializeOne(pdfDict([["Type", pdfName("Catalog")]]));
    expect(text.startsWith("%PDF-1.7\n")).toBe(true);
    expect(text).toContain("\nxref\n0 3\n");
    expect(text).toContain("0000000000 65535 f \n");
    expect(text).toContain("\ntrailer\n");
    expect(text).toContain("\nstartxref\n");
    expect(text.endsWith("%%EOF\n")).toBe(true);
  });

  test("gives every xref entry exactly twenty bytes", () => {
    const text = serializeOne(pdfDict([]));
    const start = text.indexOf("xref\n");
    const body = text.slice(text.indexOf("\n", start + "xref\n".length) + 1);
    const entries = body.slice(0, body.indexOf("trailer"));
    expect(entries.length % 20).toBe(0);
  });

  test("hashes the body into both halves of /ID", () => {
    const document = createPdfDocument();
    const root = document.add(pdfDict([]));
    const info = document.add(pdfDict([]));
    const text = latin1(document.serialize({ rootRef: root, infoRef: info, idHex: () => "ABCD" }));
    expect(text).toContain("/ID [<ABCD> <ABCD>]");
  });
});
