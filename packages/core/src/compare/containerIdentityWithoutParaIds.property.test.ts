/**
 * A comparison reads identity from the documents, not from what a save stamped.
 *
 * folio mints a `w14:paraId` for every paragraph that arrives without one, and
 * a save used to write those minted ids into the package it produced. The row
 * alignment leaned on exactly that side effect: it paired a container only when
 * its block ids were identical on both sides AND their stability had changed
 * from positional to stable, which is the shape a stamping save left behind.
 * Once a save stopped persisting a minted id, both sides read positional, that
 * path stopped firing, and one edited table cell was reported as a deleted row
 * plus an inserted row instead of an edit inside the row.
 *
 * The identity question now goes through the same owner the save asks it
 * through (`alignParagraphOrdinals`), so the two cannot come to disagree, and
 * these properties hold it there: an edit to one block is one change, in every
 * id regime a real package is in.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";
import JSZip from "jszip";
import { readFileSync } from "node:fs";
import path from "node:path";

import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";

import { FolioDocxReviewer } from "../ai-edits/headless";
import type { FolioAIBlock } from "../ai-edits/types";
import { ensureParaIds } from "../docx/ensureParaIds";
import { compareDocx } from "./compare";
import { applyEditScript, type EditScriptStep } from "./scenario";
import type { CompareChange } from "./types";

const FIXTURES_DIR = path.join(import.meta.dir, "../docx/__tests__/__fixtures__/corpus");
const OPTIONS = { author: "compare", timestamp: "2024-03-01T00:00:00.000Z" } as const;

const toArrayBuffer = (bytes: Uint8Array): ArrayBuffer =>
  bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;

/** `upstream-with-tables.docx`: a 3x3 table between two body paragraphs, no authored ids. */
const readTableFixture = (): ArrayBuffer =>
  toArrayBuffer(readFileSync(path.join(FIXTURES_DIR, "upstream-with-tables.docx")));

/** How many of the package's paragraphs carry an authored `w14:paraId`. */
const ID_COVERAGE = { none: "none", all: "all", mixed: "mixed" } as const;
type IdCoverage = (typeof ID_COVERAGE)[keyof typeof ID_COVERAGE];

const PARAGRAPH_IDS = /\sw14:(?:para|text)Id="[0-9A-Fa-f]{8}"/gu;

/**
 * Strip the ids from every other `<w:p>` of a fully stamped package, leaving
 * the `w14` namespace declared: the shape Word produces when it rewrites part
 * of a document some other producer wrote.
 */
const withAlternatingIds = async (docx: Uint8Array): Promise<ArrayBuffer> => {
  const zip = await JSZip.loadAsync(docx);
  const xml = await zip.file("word/document.xml")!.async("text");
  let ordinal = -1;
  const stripped = xml.replaceAll(/<w:p(?:\s[^>]*)?>/gu, (openTag) => {
    ordinal += 1;
    return ordinal % 2 === 0 ? openTag : openTag.replace(PARAGRAPH_IDS, "");
  });
  zip.file("word/document.xml", stripped);
  return toArrayBuffer(await zip.generateAsync({ type: "uint8array" }));
};

const baseForCoverage = async (coverage: IdCoverage): Promise<ArrayBuffer> => {
  const fixture = readTableFixture();
  switch (coverage) {
    case ID_COVERAGE.none:
      return fixture;
    case ID_COVERAGE.all:
      return toArrayBuffer((await ensureParaIds(fixture)).docx);
    case ID_COVERAGE.mixed:
      return withAlternatingIds((await ensureParaIds(fixture)).docx);
    default: {
      const unreachable: never = coverage;
      return unreachable;
    }
  }
};

const blocksOf = async (buffer: ArrayBuffer): Promise<FolioAIBlock[]> =>
  (await FolioDocxReviewer.fromBuffer(buffer)).getContent();

const changesFor = async (
  base: ArrayBuffer,
  step: EditScriptStep,
): Promise<readonly CompareChange[]> => {
  const scripted = await applyEditScript(base, [step]);
  if (scripted.isErr()) {
    throw scripted.error;
  }
  expect(scripted.value.applied).toHaveLength(1);
  const compared = await compareDocx(base, scripted.value.buffer, OPTIONS);
  if (compared.isErr()) {
    throw compared.error;
  }
  return compared.value.changes;
};

/** A row reported as deleted and reinserted is the defect, whatever the count. */
const ROW_STRUCTURE_KINDS: ReadonlySet<string> = new Set(["table-row-delete", "table-row-insert"]);

describe("one edited block is one change without authored paragraph ids", () => {
  test("the pinned counterexample: editing a table cell is not a row replacement", async () => {
    // fast-check seed 952213999 at PROPERTY_TEST_NUM_RUNS_FACTOR=15 shrank the
    // `change count never exceeds the blocks the script touched` property to
    // this script. Pinned as the scenario rather than as the seed: the script
    // is what reproduces, and it survives a change to the generators.
    const base = readTableFixture();
    const blocks = await blocksOf(base);
    // The scenario reads the fixture by index, so pin the shape it relies on.
    expect(blocks[1]?.table).toBeDefined();
    expect(blocks.every((block) => block.idStability === "positional")).toBe(true);

    const changes = await changesFor(base, {
      type: "editTableCell",
      blockIndex: 1,
      text: "AaA aaa aAa aaa",
    });

    expect(changes.map(({ kind }) => kind)).toEqual(["replace"]);
  }, 120_000);

  test(
    "editing one block reports one change in every id regime",
    async () => {
      const bases = new Map<IdCoverage, ArrayBuffer>(
        await Promise.all(
          Object.values(ID_COVERAGE).map(
            async (coverage): Promise<[IdCoverage, ArrayBuffer]> => [
              coverage,
              await baseForCoverage(coverage),
            ],
          ),
        ),
      );

      await fc.assert(
        fc.asyncProperty(
          fc.constantFrom(...Object.values(ID_COVERAGE)),
          fc.constantFrom("editTableCell" as const, "editParagraph" as const),
          fc.string({ minLength: 3, maxLength: 20, unit: "grapheme-ascii" }),
          async (coverage, kind, text) => {
            // SAFETY: every coverage was built into the map above.
            const base = bases.get(coverage)!;
            const blocks = await blocksOf(base);
            const index = blocks.findIndex((block) =>
              kind === "editTableCell" ? block.table !== undefined : block.table === undefined,
            );
            // SAFETY: the fixture holds both a table cell and a body paragraph.
            const block = blocks[index]!;
            const step: EditScriptStep =
              kind === "editTableCell"
                ? { type: "editTableCell", blockIndex: index, text: `edited ${text}` }
                : {
                    type: "replaceWords",
                    blockIndex: index,
                    // SAFETY: every block of this fixture carries text.
                    find: block.text.split(" ")[0]!,
                    replace: `edited${text}`,
                  };

            const changes = await changesFor(base, step);
            expect(changes.length).toBeLessThanOrEqual(1);
            expect(
              changes.some(({ kind: changeKind }) => ROW_STRUCTURE_KINDS.has(changeKind)),
            ).toBe(false);
          },
        ),
        propertyConfig({ numRuns: 24 }),
      );
    },
    propertyTestTimeout(180_000),
  );
});
