import { describe, expect, test } from "bun:test";
import { existsSync, readFileSync } from "node:fs";
import path from "node:path";

import {
  FOLIO_FONT_FACES,
  FOLIO_FONT_FAMILIES,
  FOLIO_FONT_FILES,
  FOLIO_FONT_PACKAGES,
  FOLIO_FONTS_STYLESHEET,
} from "./index";
import {
  createFolioFontFiles,
  folioFontFiles,
  folioFontsDirectory,
} from "./node";

const PACKAGE_ROOT = path.resolve(import.meta.dir, "..");

const WOFF_SIGNATURE = "wOFF";
const WOFF2_SIGNATURE = "wOF2";

const signature = (bytes: Uint8Array): string =>
  String.fromCharCode(...bytes.subarray(0, 4));

describe("@stll/folio-fonts", () => {
  test("every catalogued file ships, and nothing else under fonts/", () => {
    const shipped = new Set(FOLIO_FONT_FILES);
    for (const file of shipped) {
      expect(existsSync(path.join(PACKAGE_ROOT, file)), file).toBe(true);
    }
    const onDisk = [
      ...new Bun.Glob("fonts/**/*").scanSync({ cwd: PACKAGE_ROOT }),
    ].map((file) => file.split(path.sep).join("/"));
    expect(onDisk.toSorted()).toEqual([...shipped].toSorted());
  });

  test("the stylesheet declares each face once, by its .woff2 only", () => {
    const css = readFileSync(
      path.join(PACKAGE_ROOT, FOLIO_FONTS_STYLESHEET),
      "utf8",
    );
    const sources = [...css.matchAll(/url\((?<url>[^)]+)\)/gu)].map(
      (match) => match.groups?.["url"],
    );
    expect(sources).toEqual(FOLIO_FONT_FACES.map(({ woff2 }) => `./${woff2}`));
    expect(css).not.toMatch(/\.woff\)/u);
  });

  test("every family serves 400 and 700, and measures its Latin subsets", () => {
    expect(Object.keys(FOLIO_FONT_PACKAGES)).toEqual(
      FOLIO_FONT_FAMILIES.map(({ family }) => family),
    );
    for (const { family, license } of FOLIO_FONT_FAMILIES) {
      expect(["OFL-1.1", "Apache-2.0"]).toContain(license);
      const faces = FOLIO_FONT_FACES.filter((face) => face.family === family);
      for (const weight of [400, 700]) {
        for (const subset of ["latin", "latin-ext"]) {
          const face = faces.find(
            (candidate) =>
              candidate.weight === weight &&
              candidate.style === "normal" &&
              candidate.subset === subset,
          );
          expect(
            typeof face?.woff,
            `${family} ${subset} ${String(weight)}`,
          ).toBe("string");
        }
      }
    }
  });

  test("the Node reader returns the shipped bytes by @fontsource path", () => {
    const woff = folioFontFiles.read(
      "carlito",
      "files/carlito-latin-ext-700-italic.woff",
    );
    expect(woff).not.toBeNull();
    expect(signature(woff ?? new Uint8Array())).toBe(WOFF_SIGNATURE);
    const woff2 = folioFontFiles.read(
      "carlito",
      "files/carlito-greek-400-normal.woff2",
    );
    expect(signature(woff2 ?? new Uint8Array())).toBe(WOFF2_SIGNATURE);
    const ranges = folioFontFiles.read("carlito", "unicode.json");
    expect(
      JSON.parse(new TextDecoder().decode(ranges ?? new Uint8Array())),
    ).toHaveProperty("latin-ext");
  });

  test("the Node reader answers null for anything the package does not ship", () => {
    expect(
      folioFontFiles.read(
        "noto-sans-arabic",
        "files/noto-sans-arabic-latin-400-italic.woff",
      ),
    ).toBeNull();
    expect(
      folioFontFiles.read("carlito", "files/carlito-greek-400-normal.woff"),
    ).toBeNull();
    expect(folioFontFiles.read("carlito", "../../package.json")).toBeNull();
    expect(folioFontFiles.read("..", "package.json")).toBeNull();
    expect(folioFontFiles.read("helvetica", "unicode.json")).toBeNull();
  });

  test("a reader over a moved copy reads from that directory", () => {
    expect(folioFontsDirectory()).toBe(path.join(PACKAGE_ROOT, "fonts"));
    const missing = createFolioFontFiles(path.join(PACKAGE_ROOT, "absent"));
    expect(missing.read("carlito", "unicode.json")).toBeNull();
    const moved = createFolioFontFiles(`${folioFontsDirectory()}/`);
    expect(moved.read("carlito", "unicode.json")).not.toBeNull();
  });
});
