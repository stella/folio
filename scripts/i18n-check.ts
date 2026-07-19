#!/usr/bin/env bun
/**
 * Compare locale files against en.json (the source of truth).
 * Reports missing/extra keys, untranslated values (same as en), and
 * feature keys that duplicate a common.* value.
 *
 * Usage: i18n-check [langs-dir] [--sync | --write-baseline]
 *
 * langs-dir          Defaults to packages/react/src/i18n/messages.
 * --sync             Fix structural mismatches: add missing keys
 *                    (English fallback) and remove extra keys.
 * --write-baseline   Regenerate i18n-check-baseline.json, grandfathering
 *                    the current untranslated/duplicate debt so the gate
 *                    stays green while catching new regressions.
 */
import path from "node:path";

import {
  deleteNestedValue,
  getNestedValue,
  setNestedValue,
  type NestedObject,
} from "./lib/nested-object";

export type NestedMessages = NestedObject;

const flattenKeys = (obj: NestedMessages, prefix = ""): string[] => {
  const keys: string[] = [];

  for (const [key, value] of Object.entries(obj)) {
    const fullKey = prefix ? `${prefix}.${key}` : key;

    if (typeof value === "string") {
      keys.push(fullKey);
    } else {
      keys.push(...flattenKeys(value, fullKey));
    }
  }

  return keys;
};

/** Check whether all keys are sorted alphabetically (recursive). */
export const isSorted = (obj: NestedMessages): boolean => {
  const keys = Object.keys(obj);
  for (let i = 1; i < keys.length; i++) {
    const prev = keys[i - 1];
    const curr = keys[i];
    if (prev === undefined || curr === undefined) {
      continue;
    }
    if (prev > curr) {
      return false;
    }
  }
  return Object.values(obj).every((v) => typeof v === "string" || isSorted(v));
};

/** Recursively sort object keys alphabetically. */
export const sortKeys = (obj: NestedMessages): NestedMessages => {
  const sorted: NestedMessages = {};

  for (const key of Object.keys(obj).toSorted()) {
    const value = obj[key];
    if (value === undefined) {
      continue;
    }
    sorted[key] = typeof value === "string" ? value : sortKeys(value);
  }

  return sorted;
};

/**
 * Sync a locale object against the source (en) object.
 * Returns a new object with missing keys added (English
 * fallback), extra keys removed, and all keys sorted
 * alphabetically. Existing translations are preserved.
 */
export const syncMessages = (source: NestedMessages, target: NestedMessages): NestedMessages => {
  const result: NestedMessages = structuredClone(target);

  const sourceKeys = new Set(flattenKeys(source));
  const targetKeys = new Set(flattenKeys(result));

  // Delete stale keys BEFORE adding new ones: when a key changes from a leaf to
  // a namespace (source has `foo.bar`, target still has `foo`), adding `foo.bar`
  // first and then deleting the now-extra `foo` would remove the freshly-created
  // subtree, leaving the locale still missing `foo.bar`.
  for (const key of targetKeys) {
    if (!sourceKeys.has(key)) {
      deleteNestedValue(result, key);
    }
  }

  for (const key of sourceKeys) {
    if (!targetKeys.has(key)) {
      const value = getNestedValue(source, key);
      if (value !== undefined) {
        setNestedValue(result, key, value);
      }
    }
  }

  return sortKeys(result);
};

// --- value-level validation ---

/**
 * Baseline that grandfathers known debt so the gate stays green while
 * catching NEW regressions. `identicalToSource` maps a key to the locales
 * allowed to hold the English value; `duplicatesCommon` lists feature keys
 * allowed to repeat a `common.*` value; `duplicateValues` lists feature keys
 * allowed to share a value with another feature key (instead of hoisting the
 * term to `common.*`). Burn these down over time.
 */
export type CheckBaseline = {
  identicalToSource: Record<string, string[]>;
  duplicatesCommon: string[];
  duplicateValues: string[];
};

export const emptyBaseline = (): CheckBaseline => ({
  identicalToSource: {},
  duplicatesCommon: [],
  duplicateValues: [],
});

const HAS_LETTER = /\p{L}/u;

// Strings that are universally identical across languages: acronyms, format /
// standard tokens, and proper-noun product names. Extend sparingly — a common
// word that has real translations (e.g. "Free" -> "Gratis") does NOT belong
// here; scope intentional brand labels (e.g. DeepL "Free"/"Pro" tiers) per-key
// via the baseline instead.
const ALLOWED_IDENTICAL = new Set<string>([
  "OK",
  "API",
  "PDF",
  "CSV",
  "JSON",
  "HTML",
  "DOCX",
  "URL",
  "IBAN",
  "LEDES",
  "MCP",
  "OAuth",
  "SSO",
  "ID",
  "AI",
  "UI",
  "S3",
  "DeepL",
  "stella",
  "GitHub",
  "Google",
  "Microsoft",
  "Word",
  "Excel",
  "Markdown",
  "px",
  "tw",
  "in",
  "A3",
  "A4",
  "A5",
  "B5",
  "https://example.com",
  "a, b, c, ...",
  "i, ii, iii, ...",
  "A, B, C, ...",
  "I, II, III, ...",
]);

// Some localized UI terms are spelled exactly like their English source.
// Keep these semantic equivalents separate from the universal token allowlist:
// a value accepted for one locale must still be translated in every other
// locale unless it is listed there too.
const ALLOWED_IDENTICAL_BY_LOCALE = new Map<string, ReadonlySet<string>>([
  ["cs", new Set(["Executive", "Legal", "Letter", "Text"])],
  ["de", new Set(["Format", "Horizontal", "Legal", "Orange", "Position", "Text"])],
  ["es", new Set(["Color", "Diagonal", "Horizontal", "Legal", "Normal", "Vertical"])],
  ["et", new Set(["Executive", "Font"])],
  ["he", new Set(["Letter"])],
  [
    "fr",
    new Set([
      "Dimensions",
      "Double",
      "Format",
      "Horizontal",
      "Image",
      "Normal",
      "Options",
      "Orange",
      "Orientation",
      "Page",
      "Portrait",
      "Position",
      "Style",
      "Type",
      "Vertical",
    ]),
  ],
  ["hu", new Set(["Executive", "Legal", "Letter"])],
  ["lt", new Set(["Executive", "Legal", "Letter"])],
  ["lv", new Set(["Executive", "Legal", "Letter"])],
  ["pl", new Set(["Executive", "Format", "Legal", "Letter"])],
  ["pt-BR", new Set(["Diagonal", "Horizontal", "Normal", "Vertical"])],
  ["sk", new Set(["Executive", "Legal", "Letter", "Text"])],
  ["tr", new Set(["Executive", "Legal", "Letter", "Normal", "Sans Serif", "Serif"])],
]);

const KEYBOARD_SHORTCUT = /^(?:(?:Alt|Ctrl|Meta|Shift)\+)*(?:Del|[A-Z])$/u;

const stripIcuPlaceholders = (value: string): string => {
  const literal: string[] = [];
  let placeholder: string[] | null = null;

  for (const char of value) {
    if (placeholder) {
      placeholder.push(char);
      if (char === "}") {
        placeholder = null;
      }
      continue;
    }

    if (char === "{") {
      placeholder = [char];
      continue;
    }

    literal.push(char);
  }

  if (placeholder) {
    literal.push(...placeholder);
  }

  return literal.join("");
};

/** A value that is expected to read identically in every language. */
const isTriviallyIdentical = (value: string): boolean => {
  const trimmed = value.trim();
  // Ignore ICU placeholders so "{n}" / "{value}%" count as letter-free.
  const literal = stripIcuPlaceholders(trimmed);
  // Exempt only language-neutral content: no letters (numbers, punctuation,
  // placeholder-only) or an explicit allowed token. Do NOT blanket-exempt by
  // length — short words like "To"/"as" are translatable.
  return (
    !HAS_LETTER.test(literal) || ALLOWED_IDENTICAL.has(trimmed) || KEYBOARD_SHORTCUT.test(trimmed)
  );
};

const isIdenticalInLocale = (value: string, locale: string): boolean =>
  ALLOWED_IDENTICAL_BY_LOCALE.get(locale)?.has(value.trim()) ?? false;

/**
 * Keys whose locale value byte-equals the English source (untranslated),
 * excluding trivially-identical strings and baseline-grandfathered debt.
 */
export const findUntranslated = (
  source: NestedMessages,
  target: NestedMessages,
  locale: string,
  baseline: CheckBaseline,
): string[] => {
  const offenders: string[] = [];

  for (const key of flattenKeys(source)) {
    const sourceValue = getNestedValue(source, key);
    const targetValue = getNestedValue(target, key);
    if (typeof sourceValue !== "string" || sourceValue !== targetValue) {
      continue;
    }
    if (isTriviallyIdentical(sourceValue) || isIdenticalInLocale(sourceValue, locale)) {
      continue;
    }
    if (baseline.identicalToSource[key]?.includes(locale)) {
      continue;
    }
    offenders.push(key);
  }

  return offenders;
};

export type TranslationCoverage = {
  locale: string;
  total: number;
  translated: number;
  identical: number;
  approvedIdentical: number;
  missing: number;
};

type TranslationCoverageOptions = {
  source: NestedMessages;
  target: NestedMessages;
  locale: string;
  baseline: CheckBaseline;
};

/** Translation progress after excluding language-neutral and locale-identical values. */
export const getTranslationCoverage = ({
  source,
  target,
  locale,
  baseline,
}: TranslationCoverageOptions): TranslationCoverage => {
  let total = 0;
  let translated = 0;
  let identical = 0;
  let approvedIdentical = 0;
  let missing = 0;

  for (const key of flattenKeys(source)) {
    const sourceValue = getNestedValue(source, key);
    if (
      typeof sourceValue !== "string" ||
      isTriviallyIdentical(sourceValue) ||
      isIdenticalInLocale(sourceValue, locale)
    ) {
      continue;
    }

    total += 1;
    const targetValue = getNestedValue(target, key);
    if (typeof targetValue !== "string") {
      missing += 1;
      continue;
    }
    if (targetValue !== sourceValue) {
      translated += 1;
      continue;
    }

    identical += 1;
    if (baseline.identicalToSource[key]?.includes(locale)) {
      approvedIdentical += 1;
    }
  }

  return { locale, total, translated, identical, approvedIdentical, missing };
};

/** Map of each `common.*` value to the first `common.*` key that holds it. */
export const buildCommonValueMap = (source: NestedMessages): Map<string, string> => {
  const map = new Map<string, string>();

  for (const key of flattenKeys(source)) {
    if (!key.startsWith("common.")) {
      continue;
    }
    const value = getNestedValue(source, key);
    if (typeof value === "string" && !map.has(value)) {
      map.set(value, key);
    }
  }

  return map;
};

/**
 * Non-`common` en.json keys whose value duplicates an existing `common.*`
 * value (and so should reuse it). Baseline grandfathers known duplicates.
 */
export const findCommonDuplicates = (
  source: NestedMessages,
  baseline: CheckBaseline,
): { key: string; reuse: string }[] => {
  const commonByValue = buildCommonValueMap(source);
  const allow = new Set(baseline.duplicatesCommon);
  const offenders: { key: string; reuse: string }[] = [];

  for (const key of flattenKeys(source)) {
    if (key.startsWith("common.") || allow.has(key)) {
      continue;
    }
    const value = getNestedValue(source, key);
    if (typeof value !== "string") {
      continue;
    }
    const reuse = commonByValue.get(value);
    if (reuse) {
      offenders.push({ key, reuse });
    }
  }

  return offenders.toSorted((a, b) => a.key.localeCompare(b.key));
};

/**
 * Non-`common` en.json keys whose value is shared by two or more feature keys
 * in DIFFERENT top-level namespaces (and is not already a `common.*` value,
 * which `findCommonDuplicates` covers). A term repeated across feature
 * namespaces usually belongs in `common.*` so it stays consistent and is
 * translated once; this flags the duplication so it can be hoisted and reused.
 * A value repeated only within one namespace is that feature's own concern and
 * is left alone. Each offender carries the other keys it shares with. Baseline
 * grandfathers values that are intentionally duplicated.
 */
export const findSharedValueDuplicates = (
  source: NestedMessages,
  baseline: CheckBaseline,
): { key: string; shares: string[] }[] => {
  const commonByValue = buildCommonValueMap(source);
  const keysByValue = new Map<string, string[]>();

  for (const key of flattenKeys(source)) {
    if (key.startsWith("common.")) {
      continue;
    }
    const value = getNestedValue(source, key);
    if (typeof value !== "string" || isTriviallyIdentical(value) || commonByValue.has(value)) {
      continue;
    }
    const group = keysByValue.get(value);
    if (group) {
      group.push(key);
    } else {
      keysByValue.set(value, [key]);
    }
  }

  const allow = new Set(baseline.duplicateValues);
  const offenders: { key: string; shares: string[] }[] = [];

  for (const keys of keysByValue.values()) {
    if (keys.length < 2) {
      continue;
    }
    // Only flag values shared ACROSS feature namespaces (the "hoist to common.*"
    // signal). A value repeated within a single top-level namespace is that
    // feature's own business — hoisting it to common.* would not make sense — so
    // skip groups that do not span at least two namespaces.
    const namespaces = new Set(keys.map((key) => key.split(".")[0]));
    if (namespaces.size < 2) {
      continue;
    }
    const sorted = keys.toSorted((a, b) => a.localeCompare(b));
    for (const key of sorted) {
      if (allow.has(key)) {
        continue;
      }
      offenders.push({ key, shares: sorted.filter((other) => other !== key) });
    }
  }

  return offenders.toSorted((a, b) => a.key.localeCompare(b.key));
};

type StaleIdenticalApproval = { key: string; locale: string };

/** Baseline approvals that no longer describe an identical target value. */
export const findStaleIdenticalApprovals = (
  baseline: CheckBaseline,
  actualByLocale: ReadonlyMap<string, ReadonlySet<string>>,
): StaleIdenticalApproval[] => {
  const stale: StaleIdenticalApproval[] = [];

  for (const [key, locales] of Object.entries(baseline.identicalToSource)) {
    for (const locale of locales) {
      if (!actualByLocale.get(locale)?.has(key)) {
        stale.push({ key, locale });
      }
    }
  }

  return stale.toSorted(
    (left, right) => left.locale.localeCompare(right.locale) || left.key.localeCompare(right.key),
  );
};

// --- CLI ---

if (import.meta.main) {
  const args = process.argv.slice(2);
  const langsDir = args.find((a) => !a.startsWith("--")) ?? "packages/react/src/i18n/messages";
  const shouldSync = args.includes("--sync");
  const shouldWriteBaseline = args.includes("--write-baseline");

  const readLang = async (filename: string): Promise<NestedMessages> => {
    const content = await Bun.file(path.resolve(langsDir, filename)).text();
    // SAFETY: i18n JSON files conform to NestedMessages; script validates
    return JSON.parse(content) as NestedMessages;
  };

  const enRaw = await readLang("en.json");
  const enMessages = sortKeys(enRaw);
  const enKeys = new Set(flattenKeys(enMessages));

  // Baseline grandfathers existing untranslated/duplicate debt (kept beside
  // the langs dir so the *.json glob does not treat it as a locale).
  const baselinePath = path.resolve(langsDir, "..", "i18n-check-baseline.json");
  const readBaseline = async (): Promise<CheckBaseline> => {
    const file = Bun.file(baselinePath);
    if (!(await file.exists())) {
      return emptyBaseline();
    }
    // SAFETY: repo-owned baseline JSON conforms to Partial<CheckBaseline>
    const parsed = JSON.parse(await file.text()) as Partial<CheckBaseline>;
    return { ...emptyBaseline(), ...parsed };
  };
  const baseline = await readBaseline();

  const localeFiles = [...new Bun.Glob("*.json").scanSync(langsDir)]
    .filter((f) => f !== "en.json")
    .toSorted();

  // Seed/refresh the baseline from the current state, then exit.
  if (shouldWriteBaseline) {
    const identicalToSource: Record<string, string[]> = {};
    for (const file of localeFiles) {
      const locale = file.replace(/\.json$/u, "");
      // oxlint-disable-next-line no-await-in-loop -- locales must be processed in sorted order so identicalToSource accumulates deterministically
      const messages = await readLang(file);
      for (const key of findUntranslated(enMessages, messages, locale, emptyBaseline())) {
        (identicalToSource[key] ??= []).push(locale);
      }
    }
    const duplicatesCommon = findCommonDuplicates(enMessages, emptyBaseline()).map((o) => o.key);
    const duplicateValues = findSharedValueDuplicates(enMessages, emptyBaseline()).map(
      (o) => o.key,
    );
    const next: CheckBaseline = {
      duplicatesCommon: duplicatesCommon.toSorted(),
      duplicateValues: duplicateValues.toSorted(),
      identicalToSource: Object.fromEntries(
        Object.entries(identicalToSource)
          .map(([k, v]) => [k, v.toSorted()] as const)
          .toSorted(([a], [b]) => a.localeCompare(b)),
      ),
    };
    await Bun.write(baselinePath, `${JSON.stringify(next, null, 2)}\n`);
    const localeCount = Object.values(identicalToSource).reduce((a, v) => a + v.length, 0);
    console.log(
      `Wrote ${baselinePath}\n  ${Object.keys(identicalToSource).length} untranslated keys (${localeCount} locale entries), ${duplicatesCommon.length} common-duplicate keys, ${duplicateValues.length} shared-value keys grandfathered.`,
    );
    process.exit(0);
  }

  let hasIssues = false;
  const actualUntranslatedByLocale = new Map<string, ReadonlySet<string>>();
  const coverage: TranslationCoverage[] = [];

  // Check and sort en.json
  if (!isSorted(enRaw)) {
    const enPath = path.resolve(langsDir, "en.json");
    console.log(`\n${enPath}:`);
    console.log("  ~ unsorted keys");
    hasIssues = true;

    if (shouldSync) {
      await Bun.write(enPath, `${JSON.stringify(enMessages, null, 2)}\n`);
      console.log("  ✓ sorted");
    }
  } else if (shouldSync) {
    // Already sorted; nothing to write
  }

  // en.json: feature keys that duplicate a common.* value (reuse it instead).
  if (!shouldSync) {
    const duplicates = findCommonDuplicates(enMessages, baseline);
    if (duplicates.length > 0) {
      hasIssues = true;
      console.log(`\n${path.resolve(langsDir, "en.json")}:`);
      for (const { key, reuse } of duplicates) {
        console.log(`  = duplicate of ${reuse}: ${key} (reuse it, or baseline it)`);
      }
    }
  }

  // en.json: feature keys whose value is shared across feature namespaces.
  // A repeated term usually belongs in common.* so it is translated once.
  if (!shouldSync) {
    const shared = findSharedValueDuplicates(enMessages, baseline);
    if (shared.length > 0) {
      hasIssues = true;
      console.log(`\n${path.resolve(langsDir, "en.json")}:`);
      for (const { key, shares } of shared) {
        console.log(
          `  = shared value: ${key} (also ${shares.join(", ")}): consider moving the term to common.* and reusing it, or baseline it`,
        );
      }
    }
  }

  for (const file of localeFiles) {
    const locale = file.replace(/\.json$/u, "");
    // oxlint-disable-next-line no-await-in-loop -- locales reported in sorted order; console output and sync writes must stay deterministic per file
    const messages = await readLang(file);
    const langKeys = new Set(flattenKeys(messages));

    const missing = [...enKeys].filter((k) => !langKeys.has(k));
    const extra = [...langKeys].filter((k) => !enKeys.has(k));
    const unsorted = !isSorted(messages);
    const actualUntranslated = findUntranslated(enMessages, messages, locale, emptyBaseline());
    const untranslated = shouldSync ? [] : findUntranslated(enMessages, messages, locale, baseline);
    actualUntranslatedByLocale.set(locale, new Set(actualUntranslated));
    coverage.push(
      getTranslationCoverage({ source: enMessages, target: messages, locale, baseline }),
    );

    if (missing.length === 0 && extra.length === 0 && !unsorted && untranslated.length === 0) {
      continue;
    }

    hasIssues = true;
    const filePath = path.resolve(langsDir, file);
    console.log(`\n${filePath}:`);

    for (const key of missing) {
      console.log(`  + missing: ${key}`);
    }

    for (const key of extra) {
      console.log(`  - extra:   ${key}`);
    }

    for (const key of untranslated) {
      console.log(`  = untranslated (same as en): ${key}`);
    }

    if (unsorted) {
      console.log("  ~ unsorted keys");
    }

    if (shouldSync) {
      const synced = syncMessages(enMessages, messages);
      // oxlint-disable-next-line no-await-in-loop -- sequential per-locale file write paired with ordered console output
      await Bun.write(filePath, `${JSON.stringify(synced, null, 2)}\n`);
      console.log("  ✓ synced");
    }
  }

  if (!shouldSync) {
    const staleIdentical = findStaleIdenticalApprovals(baseline, actualUntranslatedByLocale);
    const actualCommonDuplicates = new Set(
      findCommonDuplicates(enMessages, emptyBaseline()).map(({ key }) => key),
    );
    const staleCommon = baseline.duplicatesCommon.filter((key) => !actualCommonDuplicates.has(key));
    const actualSharedDuplicates = new Set(
      findSharedValueDuplicates(enMessages, emptyBaseline()).map(({ key }) => key),
    );
    const staleShared = baseline.duplicateValues.filter((key) => !actualSharedDuplicates.has(key));

    if (staleIdentical.length > 0 || staleCommon.length > 0 || staleShared.length > 0) {
      hasIssues = true;
      console.log(`\n${baselinePath}:`);
      for (const { key, locale } of staleIdentical) {
        console.log(`  - stale identical approval: ${locale}: ${key}`);
      }
      for (const key of staleCommon) {
        console.log(`  - stale common-duplicate approval: ${key}`);
      }
      for (const key of staleShared) {
        console.log(`  - stale shared-value approval: ${key}`);
      }
    }

    console.log("\nTranslation coverage:");
    for (const row of coverage) {
      const percentage =
        row.total === 0 ? 100 : Math.round((row.translated / row.total) * 1000) / 10;
      const details = [
        row.identical > 0 ? `${row.identical} identical (${row.approvedIdentical} approved)` : null,
        row.missing > 0 ? `${row.missing} missing` : null,
      ].filter((detail) => detail !== null);
      const suffix = details.length > 0 ? `; ${details.join(", ")}` : "";
      console.log(
        `  ${row.locale}: ${row.translated}/${row.total} translated (${percentage}%)${suffix}`,
      );
    }
  }

  if (!hasIssues) {
    console.log("All locale files are in sync with en.json");
  } else if (!shouldSync) {
    console.log(
      "\nError: locale files or their approval baseline are out of sync with en.json.\n" +
        "Fix untranslated/duplicate findings, remove stale approvals, or explicitly approve current debt with `i18n-check <dir> --write-baseline`.",
    );
    process.exit(1);
  }
}
