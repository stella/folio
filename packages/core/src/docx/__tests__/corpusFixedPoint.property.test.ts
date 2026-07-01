/**
 * Property: parse → serialize is a fixed point over the WHOLE fixture corpus.
 *
 * The example-based `corpusRoundtrip.test.ts` asserts fidelity for a hand-picked
 * set of fixtures. This property widens that to every `.docx` under
 * `__fixtures__/corpus`: for a fast-check-selected fixture, the full editor
 * pipeline
 *
 *   parseDocx → toProseDoc → fromProseDoc → repackDocx → parseDocx
 *
 * must reach a fixed point — the re-parsed model carries the same visible body
 * text and the same paragraph count as the first parse. The nightly budget
 * (PROPERTY_TEST_NUM_RUNS_FACTOR) samples the corpus far more heavily than PR
 * CI, and any fixture added under the directory later is covered with no edit
 * here.
 *
 * Fixture provenance and licensing: see `__fixtures__/corpus/PROVENANCE.md`.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";
import { readFileSync, readdirSync } from "node:fs";
import path from "node:path";

import { propertyConfig, propertyTestTimeout } from "../../../../../test/property-testing";

import { fromProseDoc } from "../../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../../prosemirror/conversion/toProseDoc";
import type {
  BlockContent,
  Document,
  InlineSdt,
  Paragraph,
  ParagraphContent,
} from "../../types/document";
import { parseDocx } from "../parser";
import { repackDocx } from "../rezip";

const FIXTURES_DIR = path.join(import.meta.dir, "__fixtures__", "corpus");

const FIXTURE_FILES: string[] = readdirSync(FIXTURES_DIR)
  .filter((name) => name.endsWith(".docx"))
  .sort();

const readFixture = (filename: string): ArrayBuffer => {
  const bytes = readFileSync(path.join(FIXTURES_DIR, filename));
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
};

/** Collect every paragraph in the body, descending into BlockSdt wrappers. */
function collectParagraphs(blocks: readonly BlockContent[]): Paragraph[] {
  const out: Paragraph[] = [];
  for (const block of blocks) {
    if (block.type === "paragraph") {
      out.push(block);
    } else if (block.type === "blockSdt") {
      out.push(...collectParagraphs(block.content));
    }
  }
  return out;
}

/** Concatenate the text from a run/inlineSdt/hyperlink subtree. */
function textOfInline(node: ParagraphContent | InlineSdt): string {
  if (node.type === "run") {
    let s = "";
    for (const c of node.content) {
      if (c.type === "text") {
        s += c.text;
      }
    }
    return s;
  }
  if (node.type === "inlineSdt") {
    let s = "";
    for (const c of node.content) {
      s += textOfInline(c);
    }
    return s;
  }
  if (node.type === "hyperlink") {
    let s = "";
    for (const c of node.children) {
      if (c.type === "run") {
        s += textOfInline(c);
      }
    }
    return s;
  }
  return "";
}

function paragraphText(p: Paragraph): string {
  let s = "";
  for (const c of p.content) {
    s += textOfInline(c);
  }
  return s;
}

function bodyText(doc: Document): string {
  return collectParagraphs(doc.package.document.content).map(paragraphText).join("\n");
}

function paragraphCount(doc: Document): number {
  return collectParagraphs(doc.package.document.content).length;
}

async function roundTrip(original: Document): Promise<Document> {
  const back = fromProseDoc(toProseDoc(original), original);
  const repacked = await repackDocx(back, { updateModifiedDate: false });
  return parseDocx(repacked);
}

describe("corpus parse/serialize fixed point (full corpus)", () => {
  test(
    "re-parse after a repack preserves visible body text and paragraph count",
    async () => {
      await fc.assert(
        fc.asyncProperty(fc.constantFrom(...FIXTURE_FILES), async (filename) => {
          const original = await parseDocx(readFixture(filename));
          const roundTripped = await roundTrip(original);
          expect(bodyText(roundTripped)).toBe(bodyText(original));
          expect(paragraphCount(roundTripped)).toBe(paragraphCount(original));
        }),
        propertyConfig({ numRuns: 60 }),
      );
    },
    propertyTestTimeout(20_000),
  );
});
