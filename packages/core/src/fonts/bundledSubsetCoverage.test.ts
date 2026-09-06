/**
 * No code point of the diacritics fixture resolves to `.notdef`.
 *
 * `@fontsource` cuts a family into disjoint files: `latin` carries ASCII and
 * `ß` but no `ř`, `latin-ext` carries `ř` but no ASCII at all. A source that
 * serves one binary per face therefore paints empty boxes for half of a Czech,
 * Slovak or Polish sentence, and does it silently: the page still lays out,
 * every advance still comes from a real `hmtx`, and only the glyphs are gone.
 *
 * The assertion is at the resolution layer: for every code point the display
 * list paints, in the face the measurer resolved it against, *some* supplied
 * binary must return a non-zero glyph id. What a backend then does with that
 * glyph is asserted where the backend is: the PDF writer has its own gate on
 * zero glyph ids reaching a content stream.
 *
 * The font source below duplicates a few lines of
 * `scripts/bundledFontSource.ts` on purpose: a test under `src/` may not
 * import from `scripts/`, which is excluded from the build.
 */

import { afterAll, describe, expect, test } from "bun:test";
import { existsSync, readFileSync } from "node:fs";
import path from "node:path";

import { buildDisplayList } from "../display-list/build/buildDisplayList";
import type { DisplayFontFace, DisplayPrimitive } from "../display-list/types";
import { layoutDocxHeadless } from "../headless-layout";
import { getMeasureProvider, setMeasureProvider } from "../layout-engine/measure/measureProvider";
import { FONT_MAPPING } from "../utils/fontLoader";
import type { HeadlessFontRequest, HeadlessFontSource } from "./headlessMeasure";
import { installHeadlessMeasureProvider } from "./headlessMeasure";
import type { SfntFont } from "./sfnt/parse";
import { parseSfnt } from "./sfnt/parse";
import { toSfntBytes } from "./sfnt/woff";

const FONTSOURCE_DIR = path.join(
  import.meta.dir,
  "..",
  "..",
  "..",
  "react",
  "node_modules",
  "@fontsource",
);

const FIXTURE_PATH = path.join(
  import.meta.dir,
  "..",
  "docx",
  "__tests__",
  "__fixtures__",
  "corpus",
  "diacritics-latin-ext.docx",
);

/** The subsets the fixture's scripts live in, highest priority first. */
const SUBSETS = ["latin", "latin-ext"] as const;

const BOLD_WEIGHT = 700;
const REGULAR_WEIGHT = 400;

const FAMILY_DIRECTORIES = {
  Arimo: "arimo",
  Caladea: "caladea",
  Carlito: "carlito",
  Cousine: "cousine",
  Lato: "lato",
  "Noto Sans Arabic": "noto-sans-arabic",
  "Source Sans 3": "source-sans-3",
  Tinos: "tinos",
} as const;

type BundledFamily = keyof typeof FAMILY_DIRECTORIES;

const isBundledFamily = (name: string): name is BundledFamily =>
  Object.hasOwn(FAMILY_DIRECTORIES, name);

/** Word family names reach a bundled directory through `FONT_MAPPING`. */
const directoryIndex = (): ReadonlyMap<string, string> => {
  const index = new Map<string, string>();
  for (const [family, directory] of Object.entries(FAMILY_DIRECTORIES)) {
    index.set(family.toLowerCase(), directory);
  }
  for (const [authored, mapped] of Object.entries(FONT_MAPPING)) {
    if (isBundledFamily(mapped)) {
      index.set(authored.toLowerCase(), FAMILY_DIRECTORIES[mapped]);
    }
  }
  return index;
};

const DIRECTORY_INDEX = directoryIndex();

const faceFiles = ({ family, bold, italic }: HeadlessFontRequest): readonly string[] => {
  const directory = DIRECTORY_INDEX.get(family.trim().toLowerCase());
  if (directory === undefined) return [];
  const weight = bold ? BOLD_WEIGHT : REGULAR_WEIGHT;
  const style = italic ? "italic" : "normal";
  return SUBSETS.map((subset) =>
    path.join(FONTSOURCE_DIR, directory, "files", `${directory}-${subset}-${weight}-${style}.woff`),
  ).filter((filePath) => existsSync(filePath));
};

const bundledSource: HeadlessFontSource = {
  load: (request) => faceFiles(request).map((filePath) => new Uint8Array(readFileSync(filePath))),
};

const parseFace = (bytes: Uint8Array): SfntFont | null => {
  const sfnt = toSfntBytes(bytes);
  if (sfnt.isErr()) return null;
  const parsed = parseSfnt(sfnt.value);
  return parsed.isErr() ? null : parsed.value;
};

const requestFor = ({ family, weight, italic }: DisplayFontFace): HeadlessFontRequest => ({
  family,
  bold: weight >= BOLD_WEIGHT,
  italic,
});

const glyphRunsOf = (
  primitives: readonly DisplayPrimitive[],
): readonly { readonly font: number; readonly text: string }[] =>
  primitives.flatMap((primitive) => {
    switch (primitive.kind) {
      case "glyphRun":
        return [{ font: primitive.font, text: primitive.text }];
      case "clipGroup":
      case "rotateGroup":
      case "opacityGroup":
        return glyphRunsOf(primitive.children);
      case "rect":
      case "line":
      case "image":
        return [];
      default:
        primitive satisfies never;
        return [];
    }
  });

/** `family U+xxxx (c)` for every code point no supplied binary can paint. */
const unpaintableCodePoints = (
  faces: readonly DisplayFontFace[],
  runs: readonly { readonly font: number; readonly text: string }[],
): readonly string[] => {
  const parsedByFont = new Map<number, readonly SfntFont[]>();
  const reported = new Set<string>();

  for (const { font, text } of runs) {
    const face = faces[font];
    if (face === undefined) continue;
    const cached = parsedByFont.get(font);
    const fonts =
      cached ??
      bundledSource.load(requestFor(face)).flatMap((bytes) => {
        const parsed = parseFace(bytes);
        return parsed === null ? [] : [parsed];
      });
    parsedByFont.set(font, fonts);

    for (const char of text) {
      // SAFETY: iterating a string yields whole code points.
      const codePoint = char.codePointAt(0)!;
      if (fonts.some((parsed) => parsed.glyphIdFor(codePoint) !== 0)) continue;
      reported.add(`${face.family} U+${codePoint.toString(16).toUpperCase()} (${char})`);
    }
  }
  return [...reported];
};

const describeWhenBundled = existsSync(FONTSOURCE_DIR) ? describe : describe.skip;

describeWhenBundled("bundled subsets cover the diacritics fixture", () => {
  const installed = getMeasureProvider();
  afterAll(() => {
    setMeasureProvider(installed);
  });

  test("every painted code point resolves to a glyph in some served subset", async () => {
    const headless = installHeadlessMeasureProvider(bundledSource);
    const laidOut = await layoutDocxHeadless(await Bun.file(FIXTURE_PATH).arrayBuffer());
    expect(laidOut.isErr()).toBe(false);
    if (laidOut.isErr()) return;

    const list = buildDisplayList({
      layout: laidOut.value.layout,
      blockLookup: laidOut.value.blockLookup,
    });
    const runs = list.pages.flatMap((page) => glyphRunsOf(page.primitives));
    expect(runs.length).toBeGreaterThan(0);

    // A substituted face would make the coverage assertion vacuous: it would
    // hold for a font nobody chose.
    expect(headless.substitutions()).toEqual([]);
    expect(unpaintableCodePoints(list.fonts, runs)).toEqual([]);
  });
});
