import { describe, expect, test } from "bun:test";
import { collectResources, createContentStream, renderContentStream } from "./contentStream";

const render = (build: (stream: ReturnType<typeof createContentStream>) => void): string => {
  const stream = createContentStream();
  build(stream);
  const parts = stream.parts();
  return renderContentStream(parts, collectResources(parts));
};

describe("resource naming", () => {
  test("numbers names by sorted key, not by the order the painter reached them", () => {
    const text = render((stream) => {
      stream.setFont(5, 12);
      stream.setFont(1, 12);
      stream.drawImage(9);
      stream.drawImage(2);
      stream.setAlpha(0.75);
      stream.setAlpha(0.25);
    });
    expect(text).toContain("/F2 12 Tf");
    expect(text).toContain("/F1 12 Tf");
    expect(text.indexOf("/F2")).toBeLessThan(text.indexOf("/F1"));
    expect(text).toContain("/Im2 Do");
    expect(text).toContain("/Im1 Do");
    expect(text).toContain("/GS2 gs");
    expect(text).toContain("/GS1 gs");
  });

  test("gives two alphas that print alike one shared state", () => {
    const stream = createContentStream();
    stream.setAlpha(0.5);
    stream.setAlpha(0.50000001);
    const resources = collectResources(stream.parts());
    expect(resources.extGStateNames.size).toBe(1);
  });
});

describe("glyph runs", () => {
  test("groups glyphs and drops corrections that round to zero", () => {
    const text = render((stream) => {
      stream.showGlyphs([
        { glyphId: 0x41, adjustment: 0 },
        { glyphId: 0x42, adjustment: -12.5 },
        { glyphId: 0x43, adjustment: 0 },
      ]);
    });
    expect(text.trim()).toBe("[<00410042>-12.5<0043>] TJ");
  });
});

describe("operators", () => {
  test("writes colours as fractions of one", () => {
    expect(render((stream) => stream.setFillColor({ r: 255, g: 0, b: 51, a: 1 }))).toBe(
      "1 0 0.2 rg\n",
    );
  });

  test("escapes a byte string shown with a simple font", () => {
    expect(render((stream) => stream.showBytes([0x28, 0x41, 0x5c, 0xe9]))).toBe(
      "(\\(A\\\\\\351) Tj\n",
    );
  });
});
