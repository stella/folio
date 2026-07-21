import { describe, expect, test } from "bun:test";
import JSZip from "jszip";

import { docxToMarkdown } from "./docxToMarkdown";

// ── DOCX building blocks ────────────────────────────────────────────────────

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

const para = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
/** Wrap content in a block-level content control (structured document tag). */
const sdt = (inner: string) => `<w:sdt><w:sdtContent>${inner}</w:sdtContent></w:sdt>`;
const cell = (inner: string) => `<w:tc>${inner}</w:tc>`;
const row = (cells: string) => `<w:tr>${cells}</w:tr>`;
const table = (rows: string) => `<w:tbl>${rows}</w:tbl>`;

const makeDocx = async (body: string): Promise<Uint8Array> => {
  const zip = new JSZip();
  zip.file(
    "word/document.xml",
    `<w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}"><w:body>${body}</w:body></w:document>`,
  );
  return await zip.generateAsync({ type: "uint8array" });
};

/**
 * Assert that every one of `tokens` survives the DOCX -> markdown round-trip.
 * Content is dropped silently when a container type is not descended into, so
 * "every placed token appears in the output" is the invariant that catches the
 * whole class. On failure, reports exactly which tokens vanished.
 */
const expectAllExtracted = async (body: string, tokens: readonly string[]): Promise<void> => {
  const markdown = await docxToMarkdown(await makeDocx(body));
  const missing = tokens.filter((t) => !markdown.includes(t));
  expect(missing).toEqual([]);
};

// ── Fixtures: one per content-control position ──────────────────────────────

describe("docxToMarkdown content completeness — fixtures", () => {
  const fixtures: { name: string; body: string; tokens: string[] }[] = [
    {
      name: "plain body paragraph",
      body: para("PLAIN"),
      tokens: ["PLAIN"],
    },
    {
      name: "block-level content control",
      body: sdt(para("BLOCK_SDT")),
      tokens: ["BLOCK_SDT"],
    },
    {
      name: "nested content controls",
      body: sdt(sdt(para("NESTED_SDT"))),
      tokens: ["NESTED_SDT"],
    },
    {
      name: "content control inside a table cell",
      body: table(row(cell(para("CELL_PLAIN")) + cell(sdt(para("CELL_SDT"))))),
      tokens: ["CELL_PLAIN", "CELL_SDT"],
    },
    {
      name: "content control wrapping a whole cell (row level)",
      body: table(row(cell(para("ROW_A")) + sdt(cell(para("ROW_SDT_CELL"))))),
      tokens: ["ROW_A", "ROW_SDT_CELL"],
    },
    {
      name: "content control wrapping a whole row (table level)",
      body: table(sdt(row(cell(para("TBL_SDT_ROW")))) + row(cell(para("TBL_B")))),
      tokens: ["TBL_SDT_ROW", "TBL_B"],
    },
    {
      name: "content control wrapping a whole table (block level)",
      body: sdt(table(row(cell(para("SDT_TABLE"))))),
      tokens: ["SDT_TABLE"],
    },
    {
      name: "deeply nested control in a cell",
      body: table(row(cell(sdt(sdt(para("DEEP_CELL_SDT")))))),
      tokens: ["DEEP_CELL_SDT"],
    },
  ];

  test.each(fixtures)("preserves text in: $name", async ({ body, tokens }) => {
    await expectAllExtracted(body, tokens);
  });
});

// ── Fuzz: random container nestings, every token must survive ────────────────

/** Small deterministic PRNG so failures are reproducible via the seed. */
const makeRng = (seed: number) => {
  let state = (seed ^ 0x9e3779b9) >>> 0;
  return () => {
    // xorshift32
    state ^= state << 13;
    state ^= state >>> 17;
    state ^= state << 5;
    state >>>= 0;
    return state / 0xffffffff;
  };
};

describe("docxToMarkdown content completeness — fuzz", () => {
  const ITERATIONS = 250;
  const MAX_DEPTH = 4;

  for (let seed = 1; seed <= ITERATIONS; seed++) {
    test(`no content dropped for random nesting (seed ${seed})`, async () => {
      const rng = makeRng(seed);
      const tokens: string[] = [];
      const nextToken = () => {
        const t = `T${seed}x${tokens.length}z`;
        tokens.push(t);
        return t;
      };
      const pick = <T>(arr: readonly T[]): T => arr[Math.floor(rng() * arr.length)] as T;
      const times = (max: number) => 1 + Math.floor(rng() * max);

      // Paragraph content possibly wrapped in 0..2 content controls — markdown
      // treats content controls as transparent, so the text must always emerge.
      const leafParagraph = (): string => {
        let block = para(nextToken());
        for (let w = Math.floor(rng() * 3); w > 0; w--) {
          block = sdt(block);
        }
        return block;
      };

      // Cell content: paragraphs, each optionally control-wrapped. No nested
      // tables inside cells (markdown cannot represent those), which keeps the
      // invariant about *content controls* clean.
      const cellContent = (): string => {
        let inner = "";
        for (let n = times(2); n > 0; n--) {
          inner += leafParagraph();
        }
        return inner;
      };

      const buildTable = (): string => {
        const cols = times(3);
        let rows = "";
        for (let r = times(2); r > 0; r--) {
          let cells = "";
          for (let c = 0; c < cols; c++) {
            // A cell, its content, or the whole cell may be control-wrapped —
            // exercising cell-level, and (rarely) row-level sdt descent.
            cells += rng() < 0.3 ? sdt(cell(cellContent())) : cell(cellContent());
          }
          // Occasionally wrap the entire row in a content control (table level).
          rows += rng() < 0.25 ? sdt(row(cells)) : row(cells);
        }
        return table(rows);
      };

      const buildBlock = (depth: number): string => {
        if (depth <= 0 || rng() < 0.4) {
          return leafParagraph();
        }
        return pick([
          () => sdt(buildBlock(depth - 1)),
          () => buildTable(),
          () => buildBlock(depth - 1) + buildBlock(depth - 1),
        ])();
      };

      let body = "";
      for (let n = times(4); n > 0; n--) {
        body += buildBlock(MAX_DEPTH);
      }

      const markdown = await docxToMarkdown(await makeDocx(body));
      const missing = tokens.filter((t) => !markdown.includes(t));
      expect(missing).toEqual([]);
    });
  }
});
