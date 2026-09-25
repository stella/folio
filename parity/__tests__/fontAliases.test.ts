import { afterAll, describe, expect, test } from "bun:test";
import { mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import type { FontPair } from "../features";
import { reportAliasedFontFaces } from "../folioExtract";
import { planFontAliases } from "../fontAliases";
import {
  findFamilyFacesForPostscriptName,
  indexFontDirectories,
  parseFontFace,
  type FontFaceRecord,
} from "../fontFaces";
import type { PageGeom } from "../types";

type SyntheticFace = {
  family: string;
  postscriptName: string;
  typographicFamily?: string;
  weight: number;
  italic?: boolean;
};

const utf16be = (text: string): number[] =>
  [...text].flatMap((char) => [char.charCodeAt(0) >> 8, char.charCodeAt(0) & 0xff]);

/** Smallest sfnt the indexer reads: a table directory, a Windows `name`
 * table, and an `OS/2` table carrying the weight class and fsSelection. */
const buildSfnt = (face: SyntheticFace): Uint8Array => {
  const names: Array<[number, string]> = [
    [1, face.family],
    [6, face.postscriptName],
    ...(face.typographicFamily === undefined
      ? []
      : ([[16, face.typographicFamily]] as Array<[number, string]>)),
  ];
  const strings = names.map(([, value]) => utf16be(value));
  const recordsLength = names.length * 12;
  const storageOffset = 6 + recordsLength;
  const name: number[] = [0, 0, 0, names.length, storageOffset >> 8, storageOffset & 0xff];
  let stringOffset = 0;
  names.forEach(([nameId], index) => {
    const length = strings[index]?.length ?? 0;
    name.push(0, 3, 0, 1, 0x04, 0x09, 0, nameId, length >> 8, length & 0xff);
    name.push(stringOffset >> 8, stringOffset & 0xff);
    stringOffset += length;
  });
  for (const bytes of strings) name.push(...bytes);

  const os2 = Array.from({ length: 64 }, () => 0);
  os2[4] = face.weight >> 8;
  os2[5] = face.weight & 0xff;
  os2[63] = face.italic === true ? 1 : 0;

  const tables: Array<[string, number[]]> = [
    ["OS/2", os2],
    ["name", name],
  ];
  const header = [0, 1, 0, 0, 0, tables.length, 0, 0, 0, 0, 0, 0];
  let offset = 12 + tables.length * 16;
  const directory: number[] = [];
  const body: number[] = [];
  for (const [tag, data] of tables) {
    directory.push(...[...tag].map((char) => char.charCodeAt(0)));
    directory.push(0, 0, 0, 0);
    directory.push(offset >>> 24, (offset >> 16) & 0xff, (offset >> 8) & 0xff, offset & 0xff);
    directory.push(0, 0, data.length >> 8, data.length & 0xff);
    body.push(...data);
    offset += data.length;
  }
  return new Uint8Array([...header, ...directory, ...body]);
};

const readerFor =
  (bytes: Uint8Array) =>
  (offset: number, length: number): Promise<Uint8Array> =>
    Promise.resolve(bytes.slice(offset, offset + length));

const face = (
  postscriptName: string,
  family: string,
  weight: number,
  style: FontFaceRecord["style"] = "normal",
): FontFaceRecord => ({
  filePath: `/fonts/${postscriptName}.ttf`,
  postscriptName,
  family,
  weight,
  style,
});

const LOCAL_FACES = [
  face("Calibri", "Calibri", 400),
  face("Calibri-Bold", "Calibri", 700),
  face("Calibri-Italic", "Calibri", 400, "italic"),
  face("ArialMT", "Arial", 400),
];

const pair = (referenceFont: string, folioFont: string, folioRequestedFont: string): FontPair => ({
  referenceFont,
  folioFont,
  folioRequestedFont,
  referenceWidthPt: 100,
  folioWidthPt: 100,
});

describe("parseFontFace", () => {
  test("reads the PostScript name, typographic family, weight, and italic flag", async () => {
    const bytes = buildSfnt({
      family: "Example Light",
      typographicFamily: "Example",
      postscriptName: "Example-LightItalic",
      weight: 300,
      italic: true,
    });

    expect(await parseFontFace("/fonts/example.ttf", readerFor(bytes))).toEqual({
      filePath: "/fonts/example.ttf",
      postscriptName: "Example-LightItalic",
      family: "Example",
      weight: 300,
      style: "italic",
    });
  });

  test("rejects files that are not single sfnt faces", async () => {
    const collection = new Uint8Array([0x74, 0x74, 0x63, 0x66, 0, 1, 0, 0, 0, 0, 0, 1]);

    expect(await parseFontFace("/fonts/example.ttc", readerFor(collection))).toBeUndefined();
  });
});

describe("indexFontDirectories", () => {
  const directories: string[] = [];
  afterAll(async () => {
    await Promise.all(directories.map((dir) => rm(dir, { recursive: true, force: true })));
  });

  test("indexes font files and groups a PostScript name's family", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "font-index-"));
    directories.push(dir);
    await writeFile(
      path.join(dir, "example.ttf"),
      buildSfnt({ family: "Example", postscriptName: "Example", weight: 400 }),
    );
    await writeFile(
      path.join(dir, "example-bold.ttf"),
      buildSfnt({ family: "Example", postscriptName: "Example-Bold", weight: 700 }),
    );
    await writeFile(path.join(dir, "notes.txt"), "not a font");

    const faces = await indexFontDirectories([dir, path.join(dir, "missing")]);

    expect(
      findFamilyFacesForPostscriptName(faces, "ABCDEF+Example-Bold").map(
        ({ postscriptName, weight }) => [postscriptName, weight],
      ),
    ).toEqual([
      ["Example-Bold", 700],
      ["Example", 400],
    ]);
    expect(findFamilyFacesForPostscriptName(faces, "Other")).toEqual([]);
  });
});

describe("planFontAliases", () => {
  test("aliases a family Chromium could not resolve to the reference's local face", () => {
    const plan = planFontAliases(
      [
        pair("Calibri", "Arial", "Missing Sans"),
        pair("Calibri-Bold", "Arial", "Missing Sans"),
        pair("ArialMT", "Arial", "Arial"),
      ],
      LOCAL_FACES,
    );

    expect(plan.aliases).toEqual([{ requestedFamily: "Missing Sans", faceFamily: "Calibri" }]);
    expect(plan.fonts).toEqual([
      {
        family: "Missing Sans",
        filePath: "/fonts/Calibri.ttf",
        weight: 400,
        reportedFamily: "Calibri",
      },
      {
        family: "Missing Sans",
        filePath: "/fonts/Calibri-Bold.ttf",
        weight: 700,
        reportedFamily: "Calibri",
      },
      {
        family: "Missing Sans",
        filePath: "/fonts/Calibri-Italic.ttf",
        weight: 400,
        style: "italic",
        reportedFamily: "Calibri",
      },
    ]);
  });

  test("leaves families Chromium resolved, shared fallbacks, and missing faces alone", () => {
    const plan = planFontAliases(
      [
        // Chromium resolved the requested family itself.
        pair("Calibri", "Arial", "Arial"),
        // Both renderers fell back to the same family.
        pair("ArialMT", "Arial", "Shared Fallback"),
        // The reference face is not available as a local file.
        pair("AptosDisplay-Bold", "Arial", "Aptos Display"),
      ],
      LOCAL_FACES,
    );

    expect(plan).toEqual({ aliases: [], fonts: [] });
  });

  test("requires one reference face family to cover the requested family's lines", () => {
    const plan = planFontAliases(
      [
        pair("Calibri", "Arial", "Mixed"),
        pair("ArialMT", "Arial", "Mixed"),
        pair("AptosDisplay-Bold", "Arial", "Mixed"),
      ],
      LOCAL_FACES,
    );

    expect(plan.fonts).toEqual([]);
  });
});

describe("reportAliasedFontFaces", () => {
  test("reports lines painted with an aliased face under that face's family", () => {
    const pages: PageGeom[] = [
      {
        number: 1,
        widthPt: 600,
        heightPt: 800,
        lines: [
          {
            text: "Aliased",
            normText: "Aliased",
            xPt: 0,
            yPt: 0,
            widthPt: 10,
            heightPt: 10,
            fontName: "missing sans",
            requestedFontName: "Missing Sans",
            region: "body",
            direction: "ltr",
          },
          {
            text: "Native",
            normText: "Native",
            xPt: 0,
            yPt: 20,
            widthPt: 10,
            heightPt: 10,
            fontName: "Arial",
            region: "body",
            direction: "ltr",
          },
        ],
      },
    ];

    const reported = reportAliasedFontFaces(pages, [
      { family: "Missing Sans", filePath: "/fonts/Calibri.ttf", reportedFamily: "Calibri" },
    ]);

    expect(reported[0]?.lines.map(({ fontName }) => fontName)).toEqual(["Calibri", "Arial"]);
  });
});
