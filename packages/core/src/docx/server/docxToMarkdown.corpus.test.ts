import { describe, expect, test } from "bun:test";
import { readdirSync, readFileSync } from "node:fs";
import { join } from "node:path";

import { docxToMarkdown } from "./docxToMarkdown";

// Real-world DOCX corpus (content controls, tables, dropdowns, repeating
// sections, upstream-authored docs). Snapshotting the markdown extraction gives
// a reviewable diff whenever parsing or serialization changes what text or
// structure survives — a silent content drop shows up as a snapshot change.
const CORPUS_DIR = join(import.meta.dir, "../__tests__/__fixtures__/corpus");

const fixtures = readdirSync(CORPUS_DIR)
  .filter((name) => name.endsWith(".docx"))
  .sort();

describe("docxToMarkdown corpus snapshots", () => {
  test.each(fixtures)("extracts %s", async (name) => {
    const bytes = new Uint8Array(readFileSync(join(CORPUS_DIR, name)));
    const markdown = await docxToMarkdown(bytes);
    expect(markdown).toMatchSnapshot();
  });
});
