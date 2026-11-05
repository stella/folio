import { describe, expect, test } from "bun:test";

import type { PreviewDescriptor } from "../types/document";
import { sanitizeImageSrc } from "../utils/sanitizeImageSrc";
import { paintedImageSource, previewToSvgDataUrl } from "./previewSvg";

const decoded = (url: string): string =>
  decodeURIComponent(url.slice(url.indexOf(",") + 1));

const chart = (shapes: PreviewDescriptor["shapes"]): PreviewDescriptor => ({
  kind: "chart",
  extent: { width: 200, height: 100 },
  shapes,
});

describe("previewToSvgDataUrl", () => {
  test("passes the source check every painted image goes through", () => {
    const url = previewToSvgDataUrl(chart([]));
    expect(sanitizeImageSrc(url)).toBe(url);
  });

  test("draws the backdrop, then every mark in order, in the drawing's space", () => {
    const svg = decoded(
      previewToSvgDataUrl(
        chart([
          {
            type: "rect",
            x: 50,
            y: 25,
            width: 100,
            height: 50,
            color: "#102030",
          },
          {
            type: "polygon",
            color: "#405060",
            points: [
              { x: 0, y: 0 },
              { x: 200, y: 0 },
              { x: 0, y: 100 },
            ],
          },
          {
            type: "polyline",
            color: "#708090",
            strokeWidth: 2,
            points: [
              { x: 0, y: 100 },
              { x: 200, y: 0 },
            ],
          },
        ]),
      ),
    );
    expect(svg).toContain('viewBox="0 0 200 100"');
    expect(svg.match(/<(?:rect|polygon|polyline) /gu)).toEqual([
      "<rect ",
      "<rect ",
      "<polygon ",
      "<polyline ",
    ]);
    expect(svg).toContain(
      '<rect x="50" y="25" width="100" height="50" fill="rgba(16,32,48,1)"/>',
    );
    expect(svg).toContain('points="0,0 200,0 0,100"');
    expect(svg).toContain('fill="none" stroke="rgba(112,128,144,1)"');
    expect(svg).toContain('stroke-width="2"');
  });

  test("a frame with no marks is still a filled box", () => {
    const svg = decoded(previewToSvgDataUrl(chart([])));
    expect(svg.match(/<rect /gu)).toEqual(["<rect "]);
  });

  test("an extent with no area still yields a valid source", () => {
    const url = previewToSvgDataUrl({
      kind: "diagram",
      extent: { width: 0, height: 0 },
      shapes: [],
    });
    expect(decoded(url)).toContain('viewBox="0 0 1 1"');
  });

  test("a colour the descriptor carries cannot write markup", () => {
    const svg = decoded(
      previewToSvgDataUrl(
        chart([
          {
            type: "rect",
            x: 0,
            y: 0,
            width: 1,
            height: 1,
            color: '"/><script>alert(1)</script>',
          },
        ]),
      ),
    );
    expect(svg).not.toContain("script");
  });
});

describe("paintedImageSource", () => {
  test("a preview wins over the empty source that accompanies it", () => {
    const preview = chart([]);
    expect(paintedImageSource({ src: "", preview })).toBe(
      previewToSvgDataUrl(preview),
    );
  });

  test("an image with bytes keeps them", () => {
    expect(paintedImageSource({ src: "data:image/png;base64,AA==" })).toBe(
      "data:image/png;base64,AA==",
    );
  });
});
