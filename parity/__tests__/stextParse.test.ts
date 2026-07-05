import { describe, expect, test } from "bun:test";

import { parseStextXml } from "../stextParse";

// Minimal helper to build a `<char>` run for a line's reading order. The
// parser unions the non-whitespace chars' `quad`s into the line's ink bbox
// (falling back to the line `bbox` attribute when quads are absent), so the
// helper gives every char a quad spanning the whole line bbox — the union
// then equals the bbox and geometry expectations stay in one place.
const charEl = (c: string, quad: string) =>
  `<char c="${c}" quad="${quad}" x="0" y="0" bidi="0" color="#000000" alpha="#ff" flags="16"/>`;

const quadFromBbox = (bbox: string) => {
  const [x0, y0, x1, y1] = bbox.trim().split(/\s+/);
  return `${x0} ${y0} ${x1} ${y0} ${x0} ${y1} ${x1} ${y1}`;
};

const lineEl = (opts: {
  bbox: string;
  chars?: string;
  text?: string;
  font?: { name: string; size: string };
}) => {
  const textAttr = opts.text === undefined ? "" : ` text="${opts.text}"`;
  const fontOpen = opts.font
    ? `<font name="${opts.font.name}" size="${opts.font.size}">`
    : '<font name="ArialMT" size="9.12">';
  const quad = quadFromBbox(opts.bbox);
  const chars = (opts.chars ?? "")
    .split("")
    .map((c) => charEl(c, quad))
    .join("");
  return `<line bbox="${opts.bbox}" wmode="0" dir="1 0" flags="0"${textAttr}>${fontOpen}${chars}</font></line>`;
};

const pageEl = (id: number, width: string, height: string, lines: string) =>
  `<page id="page${id}" width="${width}" height="${height}"><block bbox="0 0 1 1" justify="unknown">${lines}</block></page>`;

const documentEl = (pages: string) =>
  `<?xml version="1.0"?>\n<document filename="test.pdf">\n${pages}\n</document>`;

describe("parseStextXml", () => {
  test("parses multiple pages in document order with page dimensions", () => {
    const xml = documentEl(
      [
        pageEl(1, "612", "792", lineEl({ bbox: "10 10 50 20", chars: "Hello" })),
        pageEl(2, "612", "792", lineEl({ bbox: "10 10 50 20", chars: "World" })),
      ].join("\n"),
    );

    const pages = parseStextXml(xml);

    expect(pages).toHaveLength(2);
    expect(pages[0]).toMatchObject({ number: 1, widthPt: 612, heightPt: 792 });
    expect(pages[1]).toMatchObject({ number: 2, widthPt: 612, heightPt: 792 });
    expect(pages[0]?.lines[0]?.text).toBe("Hello");
    expect(pages[1]?.lines[0]?.text).toBe("World");
  });

  test("decodes XML entities in reconstructed char text, including numeric refs", () => {
    // Reconstructed char-by-char: 'A', ' ', '&' (as &amp;), ' ', 'B', 'Ř' (hex
    // ref), 'C' (decimal ref) -> "A & BŘC".
    const chars = ["A", " ", "&amp;", " ", "B", "&#x158;", "&#67;"]
      .map((c) => charEl(c, "0 0 10 10 0 0 10 10"))
      .join("");
    const xmlLine = `<line bbox="0 0 100 10" wmode="0" dir="1 0" flags="0"><font name="ArialMT" size="9.12">${chars}</font></line>`;
    const xml = documentEl(pageEl(1, "612", "792", xmlLine));

    const pages = parseStextXml(xml);

    expect(pages[0]?.lines[0]?.text).toBe("A & BŘC");
  });

  test("drops lines whose normalized text is empty (whitespace-only or absent)", () => {
    const xml = documentEl(
      pageEl(
        1,
        "612",
        "792",
        [
          lineEl({ bbox: "0 0 10 10", chars: "   " }),
          lineEl({ bbox: "0 20 10 30" }), // no chars, no text attr
          lineEl({ bbox: "0 40 50 50", chars: "Kept" }),
        ].join(""),
      ),
    );

    const pages = parseStextXml(xml);

    expect(pages[0]?.lines).toHaveLength(1);
    expect(pages[0]?.lines[0]?.text).toBe("Kept");
  });

  test("sorts lines by (yPt, then xPt) regardless of document order", () => {
    const xml = documentEl(
      pageEl(
        1,
        "612",
        "792",
        [
          lineEl({ bbox: "50 100 90 110", chars: "SecondLineRight" }),
          lineEl({ bbox: "10 50 40 60", chars: "FirstLine" }),
          lineEl({ bbox: "10 100 40 110", chars: "SecondLineLeft" }),
        ].join(""),
      ),
    );

    const pages = parseStextXml(xml);
    const texts = pages[0]?.lines.map((l) => l.text);

    expect(texts).toEqual(["FirstLine", "SecondLineLeft", "SecondLineRight"]);
  });

  test("extracts fontName and fontSizePt from the line's first font element", () => {
    const xml = documentEl(
      pageEl(
        1,
        "612",
        "792",
        lineEl({
          bbox: "0 0 40 10",
          chars: "Text",
          font: { name: "TimesNewRomanPSMT", size: "12.5" },
        }),
      ),
    );

    const pages = parseStextXml(xml);
    const box = pages[0]?.lines[0];

    expect(box?.fontName).toBe("TimesNewRomanPSMT");
    expect(box?.fontSizePt).toBe(12.5);
  });

  test("reconstructs text from <char> elements when the line has no text attribute", () => {
    const xml = documentEl(
      pageEl(1, "612", "792", lineEl({ bbox: "0 0 60 10", chars: "NoTextAttr" })),
    );

    const pages = parseStextXml(xml);

    expect(pages[0]?.lines[0]?.text).toBe("NoTextAttr");
    expect(pages[0]?.lines[0]?.normText).toBe("NoTextAttr");
  });

  test("computes LineBox geometry (xPt/yPt/widthPt/heightPt) from the bbox attribute", () => {
    const xml = documentEl(
      pageEl(1, "612", "792", lineEl({ bbox: "12.5 20 112.5 35", chars: "X" })),
    );

    const pages = parseStextXml(xml);
    const box = pages[0]?.lines[0];

    expect(box?.xPt).toBe(12.5);
    expect(box?.yPt).toBe(20);
    expect(box?.widthPt).toBe(100);
    expect(box?.heightPt).toBe(15);
    expect(box?.region).toBe("unknown");
  });
});
