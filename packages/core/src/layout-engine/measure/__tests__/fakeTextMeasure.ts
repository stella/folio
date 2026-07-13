import { clearAllCaches } from "../cache";
import { installCanvasMeasureProvider, resetCanvasContext } from "../measureContainer";
import { getMeasureProvider, setMeasureProvider } from "../measureProvider";

/** Width contribution of a single character, given the active canvas font. */
export type FakeCharWidth = (char: string, font: string, fontKerning?: CanvasFontKerning) => number;

/** Uppercase letters render wider than everything else. */
export const uppercaseAwareCharWidth: FakeCharWidth = (char) =>
  char >= "A" && char <= "Z" ? 10 : 5;

/** Uppercase = 10, small-caps lowercase = 8, everything else = 5. */
export const smallCapsAwareCharWidth: FakeCharWidth = (char, font) => {
  if (char >= "A" && char <= "Z") {
    return 10;
  }
  if (font.includes("small-caps") && char >= "a" && char <= "z") {
    return 8;
  }
  return 5;
};

/** Every character is `px` wide, regardless of glyph. */
export const fixedCharWidth =
  (px: number): FakeCharWidth =>
  () =>
    px;

type FakeTextMeasureOptions = {
  /** Per-character width; defaults to {@link uppercaseAwareCharWidth}. */
  charWidth?: FakeCharWidth;
};

/**
 * Run `runTest` with a deterministic canvas text-measure stub installed on
 * `globalThis.document`, then restore the real document. The stub makes layout
 * measurement reproducible across machines without a real browser canvas, so
 * layout/pagination assertions stay stable. `getMeasureCount` reports how many
 * times `measureText` was invoked (for caching assertions); ignore it when not
 * needed.
 */
export function withFakeTextMeasure(
  runTest: (getMeasureCount: () => number) => void,
  options: FakeTextMeasureOptions = {},
): void {
  const charWidth = options.charWidth ?? uppercaseAwareCharWidth;
  const originalDocument = globalThis.document;
  const originalMeasureProvider = getMeasureProvider();
  let measureCount = 0;
  const fakeDocument = {
    createElement() {
      return {
        getContext() {
          return {
            font: "",
            fontKerning: "auto" as CanvasFontKerning,
            measureText(this: { font: string; fontKerning: CanvasFontKerning }, text: string) {
              measureCount += 1;
              let width = 0;
              for (const char of text) {
                width += charWidth(char, this.font, this.fontKerning);
              }
              return {
                width,
                actualBoundingBoxAscent: 8,
                actualBoundingBoxDescent: 2,
              };
            },
          };
        },
      };
    },
  } as unknown as Document;

  Object.defineProperty(globalThis, "document", {
    configurable: true,
    value: fakeDocument,
  });
  clearAllCaches();
  resetCanvasContext();
  installCanvasMeasureProvider();
  try {
    runTest(() => measureCount);
  } finally {
    setMeasureProvider(originalMeasureProvider);
    resetCanvasContext();
    clearAllCaches();
    Object.defineProperty(globalThis, "document", {
      configurable: true,
      value: originalDocument,
    });
  }
}
