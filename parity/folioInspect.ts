#!/usr/bin/env bun
/**
 * Folio DOM/page inspector for visual parity debugging.
 *
 * This complements `parity/inspect.ts`: the parity inspector compares an
 * external reference with Folio line boxes, while this command dumps Folio's
 * source-linked painted lines and spans so an agent can trace a visual line
 * back to PM positions and computed CSS without manually reading DevTools.
 */
import path from "node:path";

import { createFolioExtractor } from "./folioExtract";

type InspectFlags = {
  doc?: string | undefined;
  page: number;
  text?: string | undefined;
  maxSpans: number;
};

const usage = (): string =>
  [
    "Usage:",
    "  bun parity/folioInspect.ts --doc file.docx --page 1",
    "  bun parity/folioInspect.ts --doc file.docx --page 1 --text 'terms described'",
    "",
    "Options:",
    "  --doc <path>        DOCX to inspect",
    "  --page <n>          1-based page number (default: 1)",
    "  --text <text>       only include matching lines",
    "  --max-spans <n>     max spans per line (default: 20)",
  ].join("\n");

const parseArgs = (argv: string[]): InspectFlags => {
  const flags: InspectFlags = { page: 1, maxSpans: 20 };
  for (let i = 0; i < argv.length; i++) {
    const arg = argv[i];
    if (arg === "--doc") {
      flags.doc = argv[++i];
    } else if (arg === "--page") {
      flags.page = Number.parseInt(argv[++i] ?? "", 10);
    } else if (arg === "--text") {
      flags.text = argv[++i];
    } else if (arg === "--max-spans") {
      flags.maxSpans = Number.parseInt(argv[++i] ?? "", 10);
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
  if (!Number.isInteger(flags.maxSpans) || flags.maxSpans <= 0) {
    throw new Error("--max-spans must be a positive integer");
  }
  return flags;
};

const round = (value: number): number => Number(value.toFixed(3));

const main = async (): Promise<void> => {
  const flags = parseArgs(process.argv.slice(2));
  const doc = path.resolve(flags.doc!);
  const extractor = await createFolioExtractor();
  try {
    const page = await extractor.inspectPage(doc, flags.page);
    const needle = flags.text?.toLocaleLowerCase();
    const lines = page.lines
      .filter((line) => !needle || line.text.toLocaleLowerCase().includes(needle))
      .map((line) => ({
        index: line.index,
        text: line.text,
        region: line.region,
        rect: {
          left: round(line.rect.left),
          top: round(line.rect.top),
          width: round(line.rect.width),
          height: round(line.rect.height),
        },
        spans: line.spans.slice(0, flags.maxSpans).map((span) => ({
          text: span.text,
          className: span.className,
          pmStart: span.pmStart,
          pmEnd: span.pmEnd,
          rect: {
            left: round(span.rect.left),
            top: round(span.rect.top),
            width: round(span.rect.width),
            height: round(span.rect.height),
          },
          fontFamilyRaw: span.fontFamilyRaw,
          fontSizePx: span.fontSizePx === undefined ? undefined : round(span.fontSizePx),
          fontWeight: span.fontWeight,
          fontStyle: span.fontStyle,
          textTransform: span.textTransform,
        })),
        omittedSpans: Math.max(0, line.spans.length - flags.maxSpans),
      }));

    console.log(
      JSON.stringify(
        {
          doc,
          page: {
            pageNumber: page.pageNumber,
            domIndex: page.domIndex,
            offsetWidth: page.offsetWidth,
            offsetHeight: page.offsetHeight,
            zoomFactor: round(page.zoomFactor),
            lineCount: page.lines.length,
          },
          query: flags.text,
          lines,
        },
        null,
        2,
      ),
    );
  } finally {
    await extractor.close();
  }
};

try {
  await main();
} catch (error) {
  const err = error instanceof Error ? error : new Error(String(error));
  console.error(`folio inspect failed: ${err.message}`);
  console.error(usage());
  process.exitCode = 2;
}
