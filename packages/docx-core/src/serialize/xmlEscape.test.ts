import { describe, expect, test } from "bun:test";

import {
  escapeXmlAttribute,
  escapeXmlText,
  hasIllegalXmlCharacters,
  sanitizeXmlCharacters,
} from "./xmlEscape";

const CHAR = (code: number) => String.fromCodePoint(code);

describe("escapeXmlText", () => {
  test("escapes the five XML metacharacters", () => {
    expect(escapeXmlText(`<a href="x" id='y'>&z`)).toBe(
      "&lt;a href=&quot;x&quot; id=&apos;y&apos;&gt;&amp;z",
    );
  });

  test("escapes `]]>` so no CDATA close can form", () => {
    expect(escapeXmlText("]]>")).toBe("]]&gt;");
  });

  test("keeps tab and LF, which content preserves as written", () => {
    expect(escapeXmlText("a\tb\nc")).toBe("a\tb\nc");
  });

  test("writes CR as a reference, which §2.11 end-of-line normalisation leaves alone", () => {
    expect(escapeXmlText("a\rb")).toBe("a&#13;b");
  });
});

describe("escapeXmlAttribute", () => {
  test("escapes the five XML metacharacters", () => {
    expect(escapeXmlAttribute(`<a href="x" id='y'>&z`)).toBe(
      "&lt;a href=&quot;x&quot; id=&apos;y&apos;&gt;&amp;z",
    );
  });

  test("writes tab, LF and CR as references, which §3.3.3 would flatten to spaces", () => {
    expect(escapeXmlAttribute("a\tb\nc\rd")).toBe("a&#9;b&#10;c&#13;d");
  });
});

describe("the illegal-character contract", () => {
  const illegal = `a${CHAR(0)}b${CHAR(8)}c${CHAR(0x0b)}d${CHAR(0x0c)}e${CHAR(0x1f)}f${CHAR(
    0xff_fe,
  )}g${CHAR(0xff_ff)}h`;

  test("both escapers drop every character XML 1.0 cannot hold", () => {
    expect(escapeXmlText(illegal)).toBe("abcdefgh");
    expect(escapeXmlAttribute(illegal)).toBe("abcdefgh");
  });

  test("both escapers drop unpaired surrogate halves", () => {
    const orphaned = `a${CHAR(0xd8_00)}b${CHAR(0xdc_00)}c`;
    expect(escapeXmlText(orphaned)).toBe("abc");
    expect(escapeXmlAttribute(orphaned)).toBe("abc");
  });

  test("astral characters survive both escapers", () => {
    const page = CHAR(0x1_f4_c4);
    expect(escapeXmlText(`page ${page}`)).toBe(`page ${page}`);
    expect(escapeXmlAttribute(`page ${page}`)).toBe(`page ${page}`);
  });
});

describe("sanitizeXmlCharacters", () => {
  test("maps an unpaired surrogate to U+FFFD and drops the rest", () => {
    expect(sanitizeXmlCharacters(`a${CHAR(0xd8_00)}b${CHAR(0)}c`)).toBe("a�bc");
  });

  test("leaves a value XML can hold untouched, tab, LF and CR included", () => {
    const value = `a\tb\nc\rd${CHAR(0x1_f4_c4)}`;
    expect(sanitizeXmlCharacters(value)).toBe(value);
    expect(hasIllegalXmlCharacters(value)).toBe(false);
  });

  test("reports a value it would have to change", () => {
    expect(hasIllegalXmlCharacters(`x${CHAR(0x0b)}`)).toBe(true);
  });
});
