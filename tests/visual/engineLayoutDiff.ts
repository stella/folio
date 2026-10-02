/**
 * Pure comparison of per-document layout records taken under two browser
 * engines. Kept free of Playwright so a unit test can pin the diff rules.
 */

export type LineRecord = {
  /** First and last word of the line's text; empty when the line has none. */
  first: string;
  last: string;
  /** The measurer's claimed content width in CSS pixels, when painted. */
  width: number | null;
};

export type PageRecord = {
  /** False when the page never painted its lines (virtualised shell). */
  rendered: boolean;
  lines: LineRecord[];
};

export type FontResolution = {
  stack: string;
  weight: string;
  /** `document.fonts.check` for that stack and weight. */
  loaded: boolean;
};

export type DocumentRecord = {
  fixture: string;
  pages: PageRecord[];
  fonts: FontResolution[];
};

export type Difference =
  | { kind: "page-count"; document: string; expected: number; actual: number }
  | { kind: "line-count"; document: string; page: number; expected: number; actual: number }
  | {
      kind: "first-word" | "last-word";
      document: string;
      page: number;
      line: number;
      expected: string;
      actual: string;
      widthDelta: number | null;
    }
  | {
      kind: "width";
      document: string;
      page: number;
      line: number;
      expected: number;
      actual: number;
      widthDelta: number;
    };

/** Width differences below this are sub-pixel noise, not a layout change. */
export const WIDTH_NOISE_PX = 0.5;

const widthDeltaOf = (expected: LineRecord, actual: LineRecord): number | null =>
  expected.width === null || actual.width === null ? null : actual.width - expected.width;

/**
 * Every difference between the expected record and the actual one. Pages and
 * lines are 1-based in the output. When the line count of a page differs, lines
 * are still compared up to the shorter length so the first divergence is found.
 */
export const diffRecords = (expected: DocumentRecord, actual: DocumentRecord): Difference[] => {
  const document = expected.fixture;
  const out: Difference[] = [];
  if (expected.pages.length !== actual.pages.length) {
    out.push({
      kind: "page-count",
      document,
      expected: expected.pages.length,
      actual: actual.pages.length,
    });
  }
  const pageCount = Math.min(expected.pages.length, actual.pages.length);
  for (let pageIndex = 0; pageIndex < pageCount; pageIndex++) {
    const want = expected.pages[pageIndex];
    const got = actual.pages[pageIndex];
    if (want === undefined || got === undefined) continue;
    if (!want.rendered || !got.rendered) continue;
    const page = pageIndex + 1;
    if (want.lines.length !== got.lines.length) {
      out.push({
        kind: "line-count",
        document,
        page,
        expected: want.lines.length,
        actual: got.lines.length,
      });
    }
    const lineCount = Math.min(want.lines.length, got.lines.length);
    for (let lineIndex = 0; lineIndex < lineCount; lineIndex++) {
      const wantLine = want.lines[lineIndex];
      const gotLine = got.lines[lineIndex];
      if (wantLine === undefined || gotLine === undefined) continue;
      const line = lineIndex + 1;
      const widthDelta = widthDeltaOf(wantLine, gotLine);
      let wordsDiffer = false;
      for (const [kind, key] of [
        ["first-word", "first"],
        ["last-word", "last"],
      ] as const) {
        if (wantLine[key] === gotLine[key]) continue;
        wordsDiffer = true;
        out.push({
          kind,
          document,
          page,
          line,
          expected: wantLine[key],
          actual: gotLine[key],
          widthDelta,
        });
      }
      if (
        !wordsDiffer &&
        widthDelta !== null &&
        Math.abs(widthDelta) > WIDTH_NOISE_PX &&
        wantLine.width !== null &&
        gotLine.width !== null
      ) {
        out.push({
          kind: "width",
          document,
          page,
          line,
          expected: wantLine.width,
          actual: gotLine.width,
          widthDelta,
        });
      }
    }
  }
  return out;
};

/** One markdown table row per document, for the run summary. */
export const summarizeDocument = (fixture: string, differences: readonly Difference[]): string => {
  const kinds = new Map<string, number>();
  for (const { kind } of differences) kinds.set(kind, (kinds.get(kind) ?? 0) + 1);
  const detail =
    kinds.size === 0
      ? "-"
      : [...kinds].map(([kind, count]) => `${kind} x${String(count)}`).join(", ");
  return `| ${fixture} | ${String(differences.length)} | ${detail} |`;
};
