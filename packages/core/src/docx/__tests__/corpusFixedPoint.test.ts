/**
 * Parse → serialize is a fixed point over the WHOLE fixture corpus.
 *
 * The example-based `corpusRoundtrip.test.ts` asserts fidelity for a hand-picked
 * set of fixtures. This gate iterates EVERY `.docx` under `__fixtures__/corpus`
 * deterministically (`test.each`), so it runs on every fixture on every run
 * rather than sampling: the full editor pipeline
 *
 *   parseDocx → toProseDoc → fromProseDoc → repackDocx → parseDocx
 *
 * must reach a fixed point. Text is compared through the ProseMirror document's
 * `textContent`, which is the complete visible text (runs, SDT content,
 * hyperlinks, and any other inline text node), so the comparison cannot silently
 * skip a content shape a hand-rolled walker would miss. Structure is compared by
 * text-block count. Any fixture added under the directory later is covered with
 * no edit here.
 *
 * Fixture provenance and licensing: see `__fixtures__/corpus/PROVENANCE.md`.
 */

import { describe, expect, test } from "bun:test";
import type { Node as PMNode } from "prosemirror-model";
import { readFileSync, readdirSync } from "node:fs";
import path from "node:path";

import { fromProseDoc } from "../../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../../prosemirror/conversion/toProseDoc";
import type { Document } from "../../types/document";
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

const countTextBlocks = (doc: PMNode): number => {
  let count = 0;
  doc.descendants((node) => {
    if (node.isTextblock) {
      count += 1;
    }
    return true;
  });
  return count;
};

async function roundTrip(original: Document): Promise<Document> {
  const back = fromProseDoc(toProseDoc(original), original);
  const repacked = await repackDocx(back, { updateModifiedDate: false });
  return parseDocx(repacked);
}

describe("corpus parse/serialize fixed point (full corpus)", () => {
  // An empty / uncloned corpus must be a loud failure, not a silent pass:
  // `test.each` generates zero cases when the fixture list is empty.
  test("corpus fixtures are present", () => {
    expect(FIXTURE_FILES.length).toBeGreaterThan(0);
  });

  test.each(FIXTURE_FILES)(
    "parse → PM → repack → parse preserves visible text and block count (%s)",
    async (filename) => {
      const parsed = await parseDocx(readFixture(filename));
      const original = toProseDoc(parsed);
      const roundTripped = toProseDoc(await roundTrip(parsed));
      expect(roundTripped.textContent).toBe(original.textContent);
      expect(countTextBlocks(roundTripped)).toBe(countTextBlocks(original));
    },
  );
});
