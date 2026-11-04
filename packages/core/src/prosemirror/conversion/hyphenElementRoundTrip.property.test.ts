/**
 * `w:softHyphen` and `w:noBreakHyphen` (ECMA-376 Part 1, 17.3.3.29 and
 * 17.3.3.18) survive a save as elements, and their characters authored
 * inside `w:t` survive as text.
 *
 * The editor renders both elements as their Unicode characters, so the two
 * spellings of each look the same in ProseMirror. The save leg must still
 * write back the one the source held, wherever the run sits: a plain
 * paragraph, a hyperlink, or a tracked insertion or deletion. The expected
 * sequence comes from the generator, not from a folio read, so a loss on the
 * read side cannot cancel one on the write side.
 */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import { ZipArchive } from "@stll/docx-core/zip";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../../test/property-testing";
import { parseDocx } from "../../docx/parser";
import { createEmptyDocx, repackDocx } from "../../docx/rezip";
import type { Document } from "../../types/document";
import { fromProseDoc } from "./fromProseDoc";
import { HYPHEN_TEXT_CARRIERS } from "./hyphenTextCarriers";
import { toProseDoc } from "./toProseDoc";

setDefaultTimeout(propertyTestTimeout(30_000));

const SOFT_HYPHEN_TOKEN = "<w:softHyphen/>";
const NO_BREAK_HYPHEN_TOKEN = "<w:noBreakHyphen/>";

/** Text payloads, including both carrier characters authored as text. */
const TEXT_PAYLOADS = [
  "word",
  `pre${HYPHEN_TEXT_CARRIERS.noBreakHyphen}fix`,
  `hy${HYPHEN_TEXT_CARRIERS.softHyphen}phen`,
  HYPHEN_TEXT_CARRIERS.softHyphen,
  HYPHEN_TEXT_CARRIERS.noBreakHyphen,
] as const;

type RunChild =
  | { kind: "text"; payload: (typeof TEXT_PAYLOADS)[number] }
  | { kind: "softHyphen" }
  | { kind: "noBreakHyphen" };

const WRAPPERS = ["none", "hyperlink", "insertion", "deletion"] as const;

type RunSpec = {
  wrapper: (typeof WRAPPERS)[number];
  children: readonly RunChild[];
};

const runChildArbitrary: fc.Arbitrary<RunChild> = fc.oneof(
  fc.record({
    kind: fc.constant("text" as const),
    payload: fc.constantFrom(...TEXT_PAYLOADS),
  }),
  fc.constant({ kind: "softHyphen" as const }),
  fc.constant({ kind: "noBreakHyphen" as const }),
);

const runArbitrary: fc.Arbitrary<RunSpec> = fc.record({
  wrapper: fc.constantFrom(...WRAPPERS),
  children: fc.array(runChildArbitrary, { minLength: 1, maxLength: 4 }),
});

const paragraphArbitrary = fc.array(runArbitrary, {
  minLength: 1,
  maxLength: 5,
});

const childXml = (child: RunChild, deleted: boolean): string => {
  switch (child.kind) {
    case "text": {
      const element = deleted ? "w:delText" : "w:t";
      return `<${element} xml:space="preserve">${child.payload}</${element}>`;
    }
    case "softHyphen":
      return SOFT_HYPHEN_TOKEN;
    case "noBreakHyphen":
      return NO_BREAK_HYPHEN_TOKEN;
  }
};

const runXml = ({ wrapper, children }: RunSpec, index: number): string => {
  const run = `<w:r>${children
    .map((child) => childXml(child, wrapper === "deletion"))
    .join("")}</w:r>`;
  const revision = `w:id="${String(index + 1)}" w:author="A" w:date="2026-01-01T00:00:00Z"`;
  switch (wrapper) {
    case "none":
      return run;
    case "hyperlink":
      return `<w:hyperlink w:anchor="target">${run}</w:hyperlink>`;
    case "insertion":
      return `<w:ins ${revision}>${run}</w:ins>`;
    case "deletion":
      return `<w:del ${revision}>${run}</w:del>`;
  }
};

const documentXml = (runs: readonly RunSpec[]): string =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p>${runs
    .map(runXml)
    .join("")}</w:p><w:sectPr/></w:body></w:document>`;

const buildPackage = async (runs: readonly RunSpec[]): Promise<ArrayBuffer> => {
  const zip = await ZipArchive.read(await createEmptyDocx());
  zip.set("word/document.xml", documentXml(runs));
  return zip.generate({ compression: "store" }).then((bytes) => bytes.buffer);
};

/**
 * The paragraph as a sequence of tokens: one per text character, one per
 * hyphen element. Run and wrapper boundaries are not part of it; a save may
 * regroup runs, but it may not turn an element into a character or back.
 */
const statedTokens = (runs: readonly RunSpec[]): string[] =>
  runs.flatMap(({ children }) =>
    children.flatMap((child) =>
      child.kind === "text"
        ? [...child.payload]
        : [childXml(child, false)],
    ),
  );

const SAVED_TOKEN =
  /<w:(?:t|delText)(?:\s[^>]*)?>(?<text>[^<]*)<\/w:(?:t|delText)>|<w:softHyphen\/>|<w:noBreakHyphen\/>/g;

const savedTokens = async (saved: ArrayBuffer): Promise<string[]> => {
  const part = (await ZipArchive.read(saved)).entry("word/document.xml");
  if (!part) {
    throw new Error("the saved package has no word/document.xml");
  }
  const xml = await part.text();
  return [...xml.matchAll(SAVED_TOKEN)].flatMap((match) => {
    const text = match.groups?.["text"];
    return text === undefined ? [match[0]] : [...text];
  });
};

const repack = (document: Document): Promise<ArrayBuffer> =>
  repackDocx(document, { updateModifiedDate: false });

describe("hyphen elements and hyphen characters keep their spelling", () => {
  test("through a plain repack and through the editor", async () => {
    await assertProperty(
      fc.asyncProperty(paragraphArbitrary, async (runs) => {
        const expected = statedTokens(runs);
        const parsed = await parseDocx(await buildPackage(runs), {
          preloadFonts: false,
        });

        expect(await savedTokens(await repack(parsed))).toEqual(expected);

        const rebuilt = fromProseDoc(toProseDoc(parsed), parsed, {
          stylesheetSource: { type: "package" },
        });
        expect(await savedTokens(await repack(rebuilt))).toEqual(expected);
      }),
      {
        numRuns: 120,
        id: "hyphen elements and hyphen characters keep their spelling",
      },
    );
  });
});
