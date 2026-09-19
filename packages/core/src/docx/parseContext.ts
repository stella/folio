/**
 * The channel a parser reports a normalisation through.
 *
 * Folio accepts input Word accepts, which means normalising at the parse
 * boundary rather than refusing. Every such decision has to be visible, and
 * before this the only place that could say so was `parseDocx` itself: the
 * leaf readers are pure functions with no way to report, so a value outside
 * `ST_OnOff` or a `w:type` outside `ST_HdrFtr` was normalised in silence.
 *
 * The context is an explicit parameter, never a module-level accumulator and
 * never async-local storage. Parsers stay re-entrant, a test can hand one in
 * and read what a single reader reported, and two documents parsed at once
 * cannot write into each other's list. The cost is a threaded argument, which
 * is also the thing that makes the reporting greppable.
 */

import {
  MAX_RETAINED_PARSE_WARNINGS_PER_CODE,
  type ParseWarning,
  type ParseWarningCode,
  type ParseWarningLocation,
} from "@stll/docx-core/model";

/** What a call site states; the collector supplies the rest. */
export type ParseWarningReport = {
  code: ParseWarningCode;
  /** The element as written, prefix included. */
  element?: string;
  /** The best position this part can name, e.g. `style "Heading1"`. */
  at?: string;
  /** The value folio declined to read, as written. */
  value?: string;
  /** Text passed through from another owner, for the pass-through codes. */
  detail?: string;
  /** How many occurrences this one report stands for; defaults to 1. */
  count?: number;
};

export type ParseContext = {
  /** Record one normalisation applied to input folio accepted. */
  warn: (report: ParseWarningReport) => void;
  /** The same collector, reporting against another part or position. */
  scoped: (location: { part?: string; at?: string }) => ParseContext;
};

export type ParseWarningCollector = {
  context: ParseContext;
  /**
   * Everything recorded, in the order it was recorded, followed by one entry
   * per code whose occurrences ran past the cap. Deterministic: the same
   * document yields the same list.
   */
  warnings: () => ParseWarning[];
};

const PACKAGE_PART = "package";

export const createParseWarningCollector = (part = PACKAGE_PART): ParseWarningCollector => {
  const retained: ParseWarning[] = [];
  const retainedByCode = new Map<ParseWarningCode, number>();
  const suppressedByCode = new Map<ParseWarningCode, number>();

  const record = (location: ParseWarningLocation, report: ParseWarningReport): void => {
    const count = report.count ?? 1;
    const kept = retainedByCode.get(report.code) ?? 0;
    if (kept >= MAX_RETAINED_PARSE_WARNINGS_PER_CODE) {
      suppressedByCode.set(report.code, (suppressedByCode.get(report.code) ?? 0) + count);
      return;
    }
    retainedByCode.set(report.code, kept + 1);
    retained.push({
      code: report.code,
      location: {
        ...location,
        ...(report.element === undefined ? {} : { element: report.element }),
        ...(report.at === undefined ? {} : { at: report.at }),
      },
      ...(report.value === undefined ? {} : { value: report.value }),
      ...(report.detail === undefined ? {} : { detail: report.detail }),
      count,
    });
  };

  const contextAt = (location: ParseWarningLocation): ParseContext => ({
    warn: (report) => {
      record(location, report);
    },
    scoped: (next) => {
      // A narrower scope overrides what it states and inherits the rest, so a
      // part parser can set the part once and each element add its position.
      const at = next.at ?? location.at;
      return contextAt({
        part: next.part ?? location.part,
        ...(location.element === undefined ? {} : { element: location.element }),
        ...(at === undefined ? {} : { at }),
      });
    },
  });

  return {
    context: contextAt({ part }),
    warnings: () => [
      ...retained,
      // Sorted by code so an overflowing parse is still byte-for-byte the same
      // list twice over, whatever order the codes overflowed in.
      ...[...suppressedByCode.entries()]
        .sort(([left], [right]) => (left < right ? -1 : 1))
        .map(([code, count]) => ({
          code,
          location: { part },
          count,
          detail: `${String(count)} further occurrence(s) were counted but not retained.`,
        })),
    ],
  };
};
