import { describe, expect, test } from "bun:test";

import {
  LINE_ENDPOINT_MANIFEST_SCHEMA,
  LINE_ENDPOINT_MANIFEST_VERSION,
  compareLineEndpoints,
  createWordLineEndpointManifest,
  parseLineEndpointManifest,
} from "../lineEndpoints";
import { normalizeLineText } from "../textNorm";
import type { DocGeom, LineBox, PageGeom, Region } from "../types";

const SOURCE_SHA256 = "a".repeat(64);

describe("Word line-endpoint manifests", () => {
  test("captures normalized visual lines without local source paths", () => {
    const reference = makeGeom("word", [
      makePage(1, [
        makeLine("Clause", { xPt: 10, yPt: 20 }),
        makeLine("one", { xPt: 50, yPt: 20 }),
        makeLine("  second   line  ", { yPt: 40 }),
      ]),
    ]);

    const manifest = createWordLineEndpointManifest({
      reference,
      sourceFileName: "fixture.docx",
      sourceSha256: SOURCE_SHA256,
      capturedAt: "2026-07-14T00:00:00.000Z",
    });

    expect(manifest.schema).toBe(LINE_ENDPOINT_MANIFEST_SCHEMA);
    expect(manifest.version).toBe(LINE_ENDPOINT_MANIFEST_VERSION);
    expect(manifest.source).toEqual({ fileName: "fixture.docx", sha256: SOURCE_SHA256 });
    expect(manifest.pages).toEqual([
      {
        number: 1,
        lines: [
          { text: "Clause one", region: "body" },
          { text: "second line", region: "body" },
        ],
      },
    ]);
    expect(JSON.stringify(manifest)).not.toContain(reference.file);
  });

  test("accepts matching endpoints while ignoring geometry", () => {
    const manifest = makeManifest([
      makePage(1, [makeLine("First line"), makeLine("Second line", { yPt: 30 })]),
    ]);
    const folio = makeGeom("folio", [
      makePage(1, [
        makeLine("First line", { xPt: 120, yPt: 90, widthPt: 300 }),
        makeLine("Second line", { xPt: 90, yPt: 140, widthPt: 250 }),
      ]),
    ]);

    expect(compareLineEndpoints(manifest, folio)).toEqual({
      matches: true,
      referenceLines: 2,
      folioLines: 2,
      divergences: [],
    });
  });

  test("reports a changed line endpoint", () => {
    const manifest = makeManifest([
      makePage(1, [makeLine("Alpha beta"), makeLine("gamma delta", { yPt: 30 })]),
    ]);
    const folio = makeGeom("folio", [
      makePage(1, [makeLine("Alpha beta gamma"), makeLine("delta", { yPt: 30 })]),
    ]);

    const result = compareLineEndpoints(manifest, folio);

    expect(result.matches).toBe(false);
    expect(result.divergences).toEqual([
      {
        kind: "line-break",
        page: 1,
        referenceTexts: ["Alpha beta", "gamma delta"],
        folioTexts: ["Alpha beta gamma", "delta"],
      },
    ]);
  });

  test("reports text that moved to another page", () => {
    const manifest = makeManifest([makePage(1, [makeLine("Page edge")]), makePage(2, [])]);
    const folio = makeGeom("folio", [makePage(1, []), makePage(2, [makeLine("Page edge")])]);

    const result = compareLineEndpoints(manifest, folio);

    expect(result.matches).toBe(false);
    expect(result.divergences).toContainEqual({
      kind: "pagination",
      text: "Page edge",
      referencePage: 1,
      folioPage: 2,
    });
  });

  test("rejects unknown manifest versions and malformed hashes", () => {
    const valid = makeManifest([makePage(1, [makeLine("Text")])]);

    expect(() => parseLineEndpointManifest({ ...valid, version: 2 })).toThrow(
      "Unsupported line-endpoint manifest version: 2",
    );
    expect(() =>
      parseLineEndpointManifest({ ...valid, source: { ...valid.source, sha256: "not-a-hash" } }),
    ).toThrow("source.sha256 must be 64 lowercase hex characters");
    expect(() =>
      parseLineEndpointManifest({
        ...valid,
        source: { ...valid.source, fileName: "/private/fixture.docx" },
      }),
    ).toThrow("source.fileName must be a basename");
  });
});

const makeManifest = (pages: PageGeom[]) =>
  createWordLineEndpointManifest({
    reference: makeGeom("word", pages),
    sourceFileName: "fixture.docx",
    sourceSha256: SOURCE_SHA256,
    capturedAt: "2026-07-14T00:00:00.000Z",
  });

const makeGeom = (source: "word" | "folio", pages: PageGeom[]): DocGeom => ({
  source,
  file: source === "word" ? "/private/reference/fixture.docx" : "/workspace/fixture.docx",
  pages,
  meta:
    source === "word" ? { wordVersion: "16.99", mutool: "mutool 1.0", sha256: SOURCE_SHA256 } : {},
});

const makePage = (number: number, lines: LineBox[]): PageGeom => ({
  number,
  widthPt: 612,
  heightPt: 792,
  lines,
});

type MakeLineOptions = {
  xPt?: number;
  yPt?: number;
  widthPt?: number;
  region?: Region;
};

const makeLine = (
  text: string,
  { xPt = 10, yPt = 10, widthPt = 80, region = "body" }: MakeLineOptions = {},
): LineBox => ({
  text,
  normText: normalizeLineText(text),
  xPt,
  yPt,
  baselinePt: yPt + 9,
  widthPt,
  heightPt: 10,
  region,
});
