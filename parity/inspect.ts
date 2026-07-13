#!/usr/bin/env bun
/**
 * Deterministic parity inspector for one document/page/text.
 *
 * This is intentionally outside the HTML report path: it gives an agent a
 * small, stable JSON artifact for a single question instead of forcing manual
 * screenshot reading from the aggregate report.
 */
import path from "node:path";

import { compareGeoms, mergeVisualRows } from "./compare";
import { createFolioExtractor } from "./folioExtract";
import { getReferenceRenderer, isReferenceRendererId } from "./referenceRenderer";
import { normalizeLineText, textSimilarity } from "./textNorm";
import type { DocGeom, LineBox, PageGeom, ReferenceRendererId } from "./types";

type InspectFlags = {
  doc?: string;
  page: number;
  text?: string;
  maxPages?: number;
  ledger: boolean;
  limit: number;
  referenceId: ReferenceRendererId;
};

type Candidate = {
  rank: number;
  similarity: number;
  lineIndex: number;
  line: LineBox;
};

type LedgerRow = {
  index: number;
  referenceText?: string;
  folioText?: string;
  similarity?: number;
  referenceY?: number;
  folioY?: number;
  deltaY?: number;
  referenceX?: number;
  folioX?: number;
  deltaX?: number;
  referenceWidth?: number;
  folioWidth?: number;
  deltaWidth?: number;
};

type FolioExtractResult = Awaited<
  ReturnType<Awaited<ReturnType<typeof createFolioExtractor>>["extract"]>
>;

const usage = (): string =>
  [
    "Usage:",
    "  bun parity/inspect.ts --doc file.docx --reference libreoffice --page 1 --text 'Definitions'",
    "  bun parity/inspect.ts --doc file.docx --reference word --page 1 --ledger --max-pages 3",
    "",
    "Options:",
    "  --doc <path>        DOCX to inspect",
    "  --reference <id>    libreoffice (default) or word",
    "  --page <n>          1-based page number (default: 1)",
    "  --text <text>       line text to search for",
    "  --max-pages <n>     cap Folio extraction",
    "  --ledger            include index-by-index page ledger",
    "  --limit <n>         candidate count per side (default: 5)",
  ].join("\n");

const parseArgs = (argv: string[]): InspectFlags => {
  const flags: InspectFlags = {
    page: 1,
    ledger: false,
    limit: 5,
    referenceId: "libreoffice",
  };
  for (let i = 0; i < argv.length; i++) {
    const arg = argv[i];
    if (arg === "--doc") {
      flags.doc = argv[++i];
    } else if (arg === "--reference") {
      const value = argv[++i];
      if (!value || !isReferenceRendererId(value)) {
        throw new Error("--reference requires libreoffice or word");
      }
      flags.referenceId = value;
    } else if (arg === "--page") {
      flags.page = Number.parseInt(argv[++i] ?? "", 10);
    } else if (arg === "--text") {
      flags.text = argv[++i];
    } else if (arg === "--max-pages") {
      flags.maxPages = Number.parseInt(argv[++i] ?? "", 10);
    } else if (arg === "--ledger") {
      flags.ledger = true;
    } else if (arg === "--limit") {
      flags.limit = Number.parseInt(argv[++i] ?? "", 10);
    } else if (arg === "--help" || arg === "-h") {
      console.log(usage());
      process.exit(0);
    } else if (!flags.doc) {
      flags.doc = arg;
    } else {
      throw new Error(`Unknown argument: ${arg}`);
    }
  }
  if (!flags.doc) {
    throw new Error("--doc is required");
  }
  if (!Number.isInteger(flags.page) || flags.page <= 0) {
    throw new Error("--page must be a positive integer");
  }
  if (flags.maxPages !== undefined && (!Number.isInteger(flags.maxPages) || flags.maxPages <= 0)) {
    throw new Error("--max-pages must be a positive integer");
  }
  if (!Number.isInteger(flags.limit) || flags.limit <= 0) {
    throw new Error("--limit must be a positive integer");
  }
  return flags;
};

const pageLines = (geom: DocGeom, pageNumber: number): LineBox[] => {
  const page = geom.pages.find((candidate) => candidate.number === pageNumber);
  return page ? mergeVisualRows(page.lines) : [];
};

const pageInfo = (geom: DocGeom, pageNumber: number): PageGeom | undefined =>
  geom.pages.find((candidate) => candidate.number === pageNumber);

const candidatesFor = (lines: LineBox[], text: string, limit: number): Candidate[] => {
  const needle = normalizeLineText(text);
  return lines
    .map((line, index) => {
      const contains = line.normText.includes(needle) || needle.includes(line.normText);
      const similarity = contains ? 1 : textSimilarity(needle, line.normText);
      return { rank: 0, similarity, lineIndex: index + 1, line };
    })
    .toSorted((a, b) => b.similarity - a.similarity || a.lineIndex - b.lineIndex)
    .slice(0, limit)
    .map((candidate, index) => Object.assign({}, candidate, { rank: index + 1 }));
};

const round = (value: number | undefined): number | undefined =>
  value === undefined ? undefined : Number(value.toFixed(3));

const ledgerFor = (referenceLines: LineBox[], folioLines: LineBox[]): LedgerRow[] => {
  const max = Math.max(referenceLines.length, folioLines.length);
  const rows: LedgerRow[] = [];
  for (let i = 0; i < max; i++) {
    const reference = referenceLines[i];
    const folio = folioLines[i];
    rows.push({
      index: i + 1,
      ...(reference
        ? {
            referenceText: reference.text,
            referenceY: round(reference.yPt),
            referenceX: round(reference.xPt),
            referenceWidth: round(reference.widthPt),
          }
        : {}),
      ...(folio
        ? {
            folioText: folio.text,
            folioY: round(folio.yPt),
            folioX: round(folio.xPt),
            folioWidth: round(folio.widthPt),
          }
        : {}),
      ...(reference && folio
        ? {
            similarity: round(textSimilarity(reference.normText, folio.normText)),
            deltaY: round(folio.yPt - reference.yPt),
            deltaX: round(folio.xPt - reference.xPt),
            deltaWidth: round(folio.widthPt - reference.widthPt),
          }
        : {}),
    });
  }
  return rows;
};

const main = async (): Promise<void> => {
  const flags = parseArgs(process.argv.slice(2));
  const doc = path.resolve(flags.doc!);
  const maxPages = flags.maxPages ?? Math.max(flags.page, 1);
  const referenceRenderer = getReferenceRenderer(flags.referenceId);

  const referenceGeom = await referenceRenderer.getGeometry(doc, {});
  const extractor = await createFolioExtractor();
  let folio: FolioExtractResult;
  try {
    folio = await extractor.extract(doc, { maxPages });
  } finally {
    await extractor.close();
  }

  const limitedReferenceGeom = {
    ...referenceGeom,
    pages: referenceGeom.pages.slice(0, maxPages),
  };
  const comparison = compareGeoms(limitedReferenceGeom, folio.geom);
  const referenceLines = pageLines(referenceGeom, flags.page);
  const folioLines = pageLines(folio.geom, flags.page);
  const referencePage = pageInfo(referenceGeom, flags.page);
  const folioPage = pageInfo(folio.geom, flags.page);
  const referencePngs = await referenceRenderer.getPagePngs(doc, { maxPages });
  const counts: Record<string, number> = {};
  for (const divergence of comparison.divergences) {
    counts[divergence.kind] = (counts[divergence.kind] ?? 0) + 1;
  }

  const result = {
    doc,
    page: flags.page,
    score: comparison.score,
    counts,
    referenceRenderer: {
      id: referenceRenderer.id,
      displayName: referenceRenderer.displayName,
    },
    pages: {
      reference: referencePage && {
        widthPt: referencePage.widthPt,
        heightPt: referencePage.heightPt,
        lines: referenceLines.length,
      },
      folio: folioPage && {
        widthPt: folioPage.widthPt,
        heightPt: folioPage.heightPt,
        lines: folioLines.length,
      },
    },
    assets: {
      referencePng: referencePngs[flags.page - 1],
      folioPng: folio.screenshotPaths[flags.page - 1],
    },
    ...(flags.text
      ? {
          query: flags.text,
          referenceCandidates: candidatesFor(referenceLines, flags.text, flags.limit),
          folioCandidates: candidatesFor(folioLines, flags.text, flags.limit),
        }
      : {}),
    ...(flags.ledger ? { ledger: ledgerFor(referenceLines, folioLines) } : {}),
  };

  console.log(JSON.stringify(result, null, 2));
};

try {
  await main();
} catch (error) {
  const err = error instanceof Error ? error : new Error(String(error));
  console.error(`parity inspect failed: ${err.message}`);
  console.error(usage());
  process.exitCode = 2;
}
