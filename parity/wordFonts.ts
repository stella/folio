import path from "node:path";

import type { LocalFontDefinition } from "./folioExtract";

const WORD_FONT_DIR = "/Applications/Microsoft Word.app/Contents/Resources/DFonts";

type WordFontFace = Omit<LocalFontDefinition, "filePath"> & { fileName: string };

const APTOS_FACES = [
  { fileName: "Aptos-Light.ttf", family: "Aptos", weight: 300 },
  { fileName: "Aptos-Light-Italic.ttf", family: "Aptos", weight: 300, style: "italic" },
  { fileName: "Aptos.ttf", family: "Aptos", weight: 400 },
  { fileName: "Aptos-Italic.ttf", family: "Aptos", weight: 400, style: "italic" },
  { fileName: "Aptos-SemiBold.ttf", family: "Aptos", weight: 600 },
  { fileName: "Aptos-SemiBold-Italic.ttf", family: "Aptos", weight: 600, style: "italic" },
  { fileName: "Aptos-Bold.ttf", family: "Aptos", weight: 700 },
  { fileName: "Aptos-Bold-Italic.ttf", family: "Aptos", weight: 700, style: "italic" },
  { fileName: "Aptos-ExtraBold.ttf", family: "Aptos", weight: 800 },
  { fileName: "Aptos-ExtraBold-Italic.ttf", family: "Aptos", weight: 800, style: "italic" },
  { fileName: "Aptos-Black.ttf", family: "Aptos", weight: 900 },
  { fileName: "Aptos-Black-Italic.ttf", family: "Aptos", weight: 900, style: "italic" },
  { fileName: "Aptos-Narrow.ttf", family: "Aptos Narrow", weight: 400 },
  {
    fileName: "Aptos-Narrow-Italic.ttf",
    family: "Aptos Narrow",
    weight: 400,
    style: "italic",
  },
  { fileName: "Aptos-Narrow-Bold.ttf", family: "Aptos Narrow", weight: 700 },
  {
    fileName: "Aptos-Narrow-Bold-Italic.ttf",
    family: "Aptos Narrow",
    weight: 700,
    style: "italic",
  },
] as const satisfies ReadonlyArray<WordFontFace>;

export const wordFontDefinitions = (fontDirectory: string): LocalFontDefinition[] =>
  APTOS_FACES.map((face) => {
    const font: LocalFontDefinition = {
      family: face.family,
      filePath: path.join(fontDirectory, face.fileName),
      weight: face.weight,
    };
    if ("style" in face) {
      font.style = face.style;
    }
    return font;
  });

/** Fonts bundled with Word but not normally visible to Chromium. Missing
 * faces are skipped so an older Word installation remains usable. */
export const getAvailableWordFonts = async (): Promise<LocalFontDefinition[]> => {
  const available: LocalFontDefinition[] = [];
  for (const font of wordFontDefinitions(WORD_FONT_DIR)) {
    // oxlint-disable-next-line no-await-in-loop -- preserve manifest order for deterministic face registration
    if (await Bun.file(font.filePath).exists()) {
      available.push(font);
    }
  }
  return available;
};
