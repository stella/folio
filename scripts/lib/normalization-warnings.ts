/**
 * Every parse-boundary normaliser either warns or says why it need not.
 *
 * Folio's tolerance rule is that input Word accepts is normalised rather than
 * refused, *with a structured warning*. The failure mode is a normaliser that
 * quietly changes what the document said, and nothing about writing one makes
 * the author reach for the warning channel: it is an extra argument to thread.
 *
 * So the binding is at the file. A module named for a normalisation must
 * either reference `PARSE_WARNING_CODES` or carry an exemption naming the
 * reason it loses nothing. A compile-time alternative, a total map from a
 * normalisation kind to a code, was rejected: it needs a registry the
 * normalisers do not otherwise have, and forgetting to register a new one is
 * exactly the mistake this is meant to catch, so the map would bind only the
 * normalisers someone remembered. A check keyed on the file's existence cannot
 * be forgotten into passing.
 */

export const NORMALIZATION_MODULE_SUFFIX = "Normalization.ts";
export const PARSE_WARNING_CODES_SYMBOL = "PARSE_WARNING_CODES";
export const EXEMPTION_MARKER = "PARSE-WARNING-EXEMPT:";

export type NormalizationModule = {
  /** Path as a reader would cite it, repository-relative. */
  path: string;
  source: string;
};

export type NormalizationWarningViolation = {
  path: string;
  detail: string;
};

/** An exemption has to state a reason; a bare marker is not a decision. */
const exemptionReason = (source: string): string | undefined => {
  const marker = source.indexOf(EXEMPTION_MARKER);
  if (marker === -1) {
    return undefined;
  }
  const lineEnd = source.indexOf("\n", marker);
  return source
    .slice(marker + EXEMPTION_MARKER.length, lineEnd === -1 ? undefined : lineEnd)
    .trim();
};

export const findNormalizationWarningViolations = (
  modules: readonly NormalizationModule[],
): NormalizationWarningViolation[] => {
  const violations: NormalizationWarningViolation[] = [];
  for (const { path, source } of modules) {
    if (source.includes(PARSE_WARNING_CODES_SYMBOL)) {
      continue;
    }
    const reason = exemptionReason(source);
    if (reason === undefined) {
      violations.push({
        path,
        detail: `references no ${PARSE_WARNING_CODES_SYMBOL}; warn through the parse context, or state why it loses nothing with \`// ${EXEMPTION_MARKER} <reason>\``,
      });
      continue;
    }
    if (reason.length === 0) {
      violations.push({
        path,
        detail: `carries \`${EXEMPTION_MARKER}\` with no reason after it`,
      });
    }
  }
  return violations;
};
