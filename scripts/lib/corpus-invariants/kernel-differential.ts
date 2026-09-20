/**
 * Read every package twice, with two independent implementations.
 *
 * The Rust kernel in `crates/docx-kernel` and folio's TypeScript parser open
 * the same ZIP, resolve the same namespaces and answer the same questions
 * about it. Neither consults the other, so a disagreement is a defect in one
 * of them and no amount of round-tripping either one against itself would
 * show it.
 *
 * What the two can be *asked* is narrower than what they both compute, and the
 * comparisons below are only the ones whose definitions provably coincide. The
 * rest are dropped on purpose: a differential that fires on a definitional
 * difference teaches the reader to ignore it, which costs more than the
 * comparison was worth.
 *
 * ## What the kernel counts as a paragraph
 *
 * Established by reading `crates/docx-kernel/src/projection/ooxml.rs` and its
 * tests, because the TypeScript walk has to mean the same thing:
 *
 * - The main document part only. `extract_document_parts` resolves the
 *   document, styles and numbering parts and nothing else, so no header,
 *   footer, footnote, endnote or glossary paragraph is ever projected.
 * - Paragraphs inside tables, at their document-order position among body
 *   paragraphs; a cell paragraph is not deferred to the end.
 * - Paragraphs inside `w:sdt`, which is structurally transparent.
 * - Empty paragraphs. `<w:p/>` projects with empty text.
 * - *Not* paragraphs inside `w:txbxContent`: the textbox frame swallows every
 *   descendant, so a shape's or a VML box's text contributes nothing.
 *
 * The TypeScript walk therefore descends through block content, table cells
 * and block content controls, and never enters a run: shape text bodies are
 * the only way back from inline to block content, and the kernel does not
 * follow it either.
 *
 * ## Why the count itself is not reported
 *
 * Two constructs make the counts differ without either side being wrong:
 *
 * - A deleted paragraph mark. The kernel merges a paragraph whose mark
 *   carries `w:del` into its successor and renumbers, which is what a reader
 *   accepting the deletion would see; folio keeps both paragraphs and the
 *   mark. Every reviewed document hits this.
 * - `mc:AlternateContent`. The kernel evaluates `mc:Requires` and takes the
 *   fallback when the choice needs a namespace it does not support; folio
 *   takes the first `mc:Choice` unconditionally. Word writes this pair around
 *   every modern shape.
 *
 * So the counts are compared, but only as the precondition for everything
 * that is keyed by ordinal. When they differ the two sides are describing
 * different documents and no per-paragraph claim survives; the file is
 * reported as nothing rather than as a defect.
 */

import { readFile } from "node:fs/promises";

// By path, not by specifier: the gate's scripts are not a workspace package,
// and `@stll/docx-core` is reachable only from the packages that depend on it.
import {
  type DocxProjectionWire,
  initializeDocxProjection,
  projectCompressedDocx,
} from "../../../packages/docx-core/src/projection";
import type { BlockContent } from "@stll/folio-core/types/document";
import { Result } from "better-result";

import { type CorpusFailure, failureFromAssertion, failureFromError } from "../corpus-signature";
import {
  type CorpusInvariantInput,
  type CorpusInvariantOutcome,
  EXTENDED_CORPUS_INVARIANTS,
  timeStage,
} from "./contract";

const WASM_URL = new URL(
  "../../../packages/docx-core/src/generated/docx_kernel_bg.wasm",
  import.meta.url,
);

let runtime: Promise<void> | undefined;

/**
 * Load the checked-in kernel at first use, not at import.
 *
 * Bun resolves no asset by URL, so the runtime needs the bytes handed to it.
 * The package memoizes its own initialization; this memoizes the read, and
 * clears itself on failure so a worker that lost the file once is not stuck
 * replaying the same rejected promise for every remaining file.
 */
const readyKernel = (): Promise<void> => {
  runtime ??= readFile(WASM_URL)
    .then((wasm) => initializeDocxProjection({ wasm }))
    .catch((cause: unknown) => {
      runtime = undefined;
      throw cause;
    });
  return runtime;
};

/** What the TypeScript model says about one paragraph, in the kernel's terms. */
type TypeScriptParagraphFacts = {
  /** The paragraph's own `w:pStyle`, which is also all the kernel reports. */
  styleId: string | null;
  /** The paragraph's own `w:outlineLvl`, never a style's. */
  outlineLevel: number | null;
};

/** What the TypeScript model says about one table's shape. */
type TypeScriptTableFacts = {
  /** Cell count per row, in row order. */
  cellsPerRow: readonly number[];
};

export type TypeScriptDocumentFacts = {
  paragraphs: readonly TypeScriptParagraphFacts[];
  tables: readonly TypeScriptTableFacts[];
};

type FactAccumulator = {
  paragraphs: TypeScriptParagraphFacts[];
  tables: TypeScriptTableFacts[];
};

const collectBlocks = (blocks: readonly BlockContent[], into: FactAccumulator): void => {
  for (const block of blocks) {
    switch (block.type) {
      case "paragraph": {
        into.paragraphs.push({
          styleId: block.formatting?.styleId ?? null,
          outlineLevel: block.formatting?.outlineLevel ?? null,
        });
        break;
      }
      case "table": {
        // Pushed before its contents, so table ordinals follow `w:tbl` start
        // order, which is how the kernel's counter runs.
        into.tables.push({ cellsPerRow: block.rows.map((row) => row.cells.length) });
        for (const row of block.rows) {
          for (const cell of row.cells) {
            collectBlocks(cell.content, into);
          }
        }
        break;
      }
      case "blockSdt": {
        collectBlocks(block.content, into);
        break;
      }
      // Opaque markup, and a range marker: no paragraph and no table for the
      // kernel to count.
      case "preservedBlock":
      case "bookmarkStart":
      case "bookmarkEnd":
        break;
      default: {
        const unreachable: never = block;
        return unreachable;
      }
    }
  }
};

/** The TypeScript half of the comparison, read from the main document part. */
export const typeScriptDocumentFacts = (
  blocks: readonly BlockContent[],
): TypeScriptDocumentFacts => {
  const accumulator: FactAccumulator = { paragraphs: [], tables: [] };
  collectBlocks(blocks, accumulator);
  return accumulator;
};

/**
 * Messages carry no per-file particulars and no numbers.
 *
 * A signature erases digit runs before it is counted, so a message that named
 * the two outline levels would read as `TS says N, kernel says N` and collapse
 * into the same signature as every other outline disagreement. The direction
 * and the classification are what survive normalization, so they carry the
 * whole finding.
 */
const STYLE_ID_DISAGREES = "per-paragraph style id disagrees";
const STYLE_ID_KERNEL_NONE =
  "per-paragraph style id: the kernel reports none where TypeScript reports one";
const STYLE_ID_TYPESCRIPT_NONE =
  "per-paragraph style id: TypeScript reports none where the kernel reports one";
const OUTLINE_LEVEL_DISAGREES = "direct outline level disagrees";
const OUTLINE_LEVEL_KERNEL_NONE =
  "direct outline level: the kernel reports none where TypeScript reports one";
const TABLE_MISSING = "table structure: the kernel reports a table the TypeScript model lacks";
const TABLE_ROW_MISSING = "table structure: the kernel reports a row the TypeScript model lacks";
const TABLE_CELL_MISSING = "table structure: the kernel reports a cell the TypeScript model lacks";
const TABLE_ID_SHAPE = "kernel table identifier is not the projected shape";

const TABLE_ID_RE = /^table-(\d+)$/u;

const KERNEL_STYLE_ID = 5;
const KERNEL_STRUCTURE = 4;
const OUTLINE_LEVEL_FAMILY = 4;

const compareStyleIds = (
  projection: DocxProjectionWire,
  facts: TypeScriptDocumentFacts,
  found: Set<string>,
): void => {
  projection[1].forEach((paragraph, ordinal) => {
    const kernelStyleId = paragraph[KERNEL_STYLE_ID];
    // SAFETY: the caller proved the two paragraph lists are the same length.
    const { styleId } = facts.paragraphs[ordinal] as TypeScriptParagraphFacts;
    if (kernelStyleId === styleId) {
      return;
    }
    if (kernelStyleId === null) {
      found.add(STYLE_ID_KERNEL_NONE);
      return;
    }
    found.add(styleId === null ? STYLE_ID_TYPESCRIPT_NONE : STYLE_ID_DISAGREES);
  });
};

/**
 * Only where folio holds a direct `w:outlineLvl`.
 *
 * The kernel's fact is the resolved level: document defaults, then the
 * `w:basedOn` chain, then the paragraph's own value. folio's model carries
 * only the paragraph's own value and leaves resolution to the editor layer,
 * so the two agree exactly on the paragraphs that have a direct value, where
 * the kernel's last tier wins, and say different things everywhere else. A
 * paragraph the kernel gives a level and folio does not is that expected
 * inheritance, not a defect.
 */
const compareOutlineLevels = (
  projection: DocxProjectionWire,
  facts: TypeScriptDocumentFacts,
  found: Set<string>,
): void => {
  const family = projection[2][OUTLINE_LEVEL_FAMILY];
  // `unknown` is the kernel declining to claim completeness, not a disagreement.
  if (family[0] !== "known") {
    return;
  }
  const kernelLevels = new Map(family[1].map(([ordinal, level]) => [ordinal, level]));
  facts.paragraphs.forEach(({ outlineLevel }, ordinal) => {
    if (outlineLevel === null) {
      return;
    }
    const kernelLevel = kernelLevels.get(ordinal);
    if (kernelLevel === undefined) {
      found.add(OUTLINE_LEVEL_KERNEL_NONE);
      return;
    }
    if (kernelLevel !== outlineLevel) {
      found.add(OUTLINE_LEVEL_DISAGREES);
    }
  });
};

/**
 * One direction only: the kernel may not name a coordinate folio lacks.
 *
 * The kernel reports table structure per paragraph, so it observes a row or a
 * column only where a paragraph sits directly in it. A cell whose whole
 * content is a nested table contributes nothing, which makes the kernel's
 * observed extent a lower bound on the table's real shape. Requiring equality
 * would fire on that; requiring containment still catches the defect worth
 * catching, which is folio dropping a table, a row or a cell the kernel saw.
 */
const compareTableStructure = (
  projection: DocxProjectionWire,
  facts: TypeScriptDocumentFacts,
  found: Set<string>,
): void => {
  for (const paragraph of projection[1]) {
    const structure = paragraph[KERNEL_STRUCTURE];
    if (structure.length === 0) {
      continue;
    }
    const [, tableId, row, column] = structure;
    const ordinal = TABLE_ID_RE.exec(tableId)?.at(1);
    if (ordinal === undefined) {
      found.add(TABLE_ID_SHAPE);
      return;
    }
    const table = facts.tables.at(Number(ordinal));
    if (table === undefined) {
      found.add(TABLE_MISSING);
      continue;
    }
    const cells = table.cellsPerRow.at(row);
    if (cells === undefined) {
      found.add(TABLE_ROW_MISSING);
      continue;
    }
    if (column >= cells) {
      found.add(TABLE_CELL_MISSING);
    }
  }
};

/**
 * Every disagreement the two implementations can be held to.
 *
 * Pure, so the comparison is testable against hand-built wire tuples without
 * loading WebAssembly. Each message is emitted at most once: a style mapping
 * that broke on one paragraph broke on all of them, and one defect must not
 * arrive as four hundred failures.
 */
export const kernelDifferentialFailures = (
  projection: DocxProjectionWire,
  facts: TypeScriptDocumentFacts,
): CorpusFailure[] => {
  if (projection[1].length !== facts.paragraphs.length) {
    return [];
  }
  const found = new Set<string>();
  compareStyleIds(projection, facts, found);
  compareOutlineLevels(projection, facts, found);
  compareTableStructure(projection, facts, found);
  return [...found].map((message) =>
    failureFromAssertion(EXTENDED_CORPUS_INVARIANTS.kernelDifferential, message),
  );
};

export const runKernelDifferentialInvariant = async ({
  bytes,
  parsed,
}: CorpusInvariantInput): Promise<CorpusInvariantOutcome> => {
  const timings: Record<string, number> = {};

  const projected = await timeStage(timings, "kernel-project", () =>
    Result.tryPromise({
      try: async () => {
        await readyKernel();
        return projectCompressedDocx(bytes);
      },
      catch: (cause: unknown) => cause,
    }),
  );
  // A kernel that refuses a package folio accepted is the finding, not an outage.
  if (projected.isErr()) {
    return {
      failures: [failureFromError(EXTENDED_CORPUS_INVARIANTS.kernelDifferential, projected.error)],
      timings,
    };
  }

  const failures = await timeStage(timings, "compare", () =>
    Promise.resolve(
      kernelDifferentialFailures(
        projected.value,
        typeScriptDocumentFacts(parsed.package.document.content),
      ),
    ),
  );
  return { failures, timings };
};
