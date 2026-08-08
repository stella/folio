/**
 * Generate the Unicode Joining_Type tables the cursive-shaping helpers read.
 *
 * Joining_Type decides whether two adjacent characters were cursively connected
 * before a run boundary split them. Hand-listing those classes is a mirror that
 * drifts: the Arabic block alone carries 615 dual-joining and 153 right-joining
 * entries, and the great majority of Transparent characters are not listed in
 * `ArabicShaping.txt` at all — they inherit it from General_Category Mn/Me/Cf.
 * A hand-maintained table would silently misclassify every harakat.
 *
 * So the table is derived from the pinned UCD files and committed as generated
 * output, with `check` wired into CI so a Unicode bump cannot land half-applied.
 *
 * Usage:
 *   bun scripts/generate-joining-types.ts write
 *   bun scripts/generate-joining-types.ts check
 */

import { createHash } from "node:crypto";
import { readFile, writeFile } from "node:fs/promises";
import path from "node:path";

import { TaggedError } from "better-result";

const REPO_ROOT = path.resolve(import.meta.dir, "..");
const OUTPUT_PATH = path.join(REPO_ROOT, "packages/core/src/utils/joiningTypes.gen.ts");

/** UCD sources, resolved from `specifications/sources.json` cache paths. */
const ARABIC_SHAPING_ID = "unicode-arabic-shaping";
const GENERAL_CATEGORY_ID = "unicode-derived-general-category";

class GenerateJoiningTypesError extends TaggedError("GenerateJoiningTypesError")<{
  message: string;
}>() {}

/**
 * Joining_Type values this generator distinguishes.
 *
 * `C` (Join_Causing, e.g. ZWJ) behaves as joining on both sides, so it folds
 * into both direction tables rather than getting one of its own.
 */
type JoiningType = "C" | "D" | "L" | "R" | "T" | "U";

type Range = { start: number; end: number };

type SourceEntry = { id: string; cachePath: string; version: string; sha256: string };

const readSourceEntries = async (): Promise<Map<string, SourceEntry>> => {
  const raw = await readFile(path.join(REPO_ROOT, "specifications/sources.json"), "utf8");
  const parsed: unknown = JSON.parse(raw);
  if (
    typeof parsed !== "object" ||
    parsed === null ||
    !("sources" in parsed) ||
    !Array.isArray(parsed.sources)
  ) {
    throw new GenerateJoiningTypesError({
      message: "specifications/sources.json has no `sources` array",
    });
  }
  const entries = new Map<string, SourceEntry>();
  for (const source of parsed.sources) {
    if (
      typeof source === "object" &&
      source !== null &&
      "id" in source &&
      "cachePath" in source &&
      "version" in source &&
      "sha256" in source &&
      typeof source.id === "string" &&
      typeof source.cachePath === "string" &&
      typeof source.version === "string" &&
      typeof source.sha256 === "string"
    ) {
      entries.set(source.id, {
        id: source.id,
        cachePath: source.cachePath,
        version: source.version,
        sha256: source.sha256,
      });
    }
  }
  return entries;
};

const requireSource = (entries: Map<string, SourceEntry>, id: string): SourceEntry => {
  const entry = entries.get(id);
  if (!entry) {
    throw new GenerateJoiningTypesError({
      message: `specifications/sources.json is missing the \`${id}\` source. Add it, then run \`bun run specifications:fetch\`.`,
    });
  }
  return entry;
};

const readCachedSource = async (entry: SourceEntry): Promise<string> => {
  const absolute = path.join(REPO_ROOT, entry.cachePath);
  let contents: string;
  try {
    contents = await readFile(absolute, "utf8");
  } catch {
    throw new GenerateJoiningTypesError({
      message: `Cached source missing for \`${entry.id}\` at ${entry.cachePath}. Run \`bun run specifications:fetch\` first.`,
    });
  }
  // Verify against the manifest digest. Without this an edited cache file would
  // make `write` emit a table nobody can reproduce while `check` still passed,
  // because both sides would read the same tampered input.
  const actual = digest(contents);
  if (actual !== entry.sha256) {
    throw new GenerateJoiningTypesError({
      message: `Cached source \`${entry.id}\` does not match its manifest digest (expected ${entry.sha256}, read ${actual}). Delete ${entry.cachePath} and re-run \`bun run specifications:fetch\`.`,
    });
  }
  return contents;
};

const digest = (text: string): string => createHash("sha256").update(text).digest("hex");

/** Strip a UCD comment tail and surrounding whitespace from one line. */
const dataPart = (line: string): string => {
  const hash = line.indexOf("#");
  return (hash === -1 ? line : line.slice(0, hash)).trim();
};

/**
 * Parse `ArabicShaping.txt`: `code; NAME; JoiningType; JoiningGroup`.
 *
 * Every entry is a single code point; the file carries no ranges.
 */
const parseArabicShaping = (text: string): Map<number, JoiningType> => {
  const explicit = new Map<number, JoiningType>();
  for (const line of text.split("\n")) {
    const data = dataPart(line);
    if (data.length === 0) continue;
    const fields = data.split(";").map((f) => f.trim());
    const [code, , joining] = fields;
    if (code === undefined || joining === undefined) continue;
    if (!/^[CDLRTU]$/u.test(joining)) {
      throw new GenerateJoiningTypesError({
        message: `Unexpected Joining_Type \`${joining}\` on line: ${line}`,
      });
    }
    explicit.set(Number.parseInt(code, 16), joining as JoiningType);
  }
  if (explicit.size === 0) {
    throw new GenerateJoiningTypesError({
      message: "ArabicShaping.txt yielded no entries — parser or source is wrong",
    });
  }
  return explicit;
};

/**
 * Collect the Mn/Me/Cf ranges from `DerivedGeneralCategory.txt`.
 *
 * These default to Joining_Type Transparent, which is where nearly all
 * transparent characters come from — combining marks sit between a letter and
 * its neighbour without breaking the cursive connection.
 */
const parseTransparentDefaults = (text: string): Range[] => {
  const ranges: Range[] = [];
  for (const line of text.split("\n")) {
    const data = dataPart(line);
    if (data.length === 0) continue;
    const [codes, category] = data.split(";").map((f) => f.trim());
    if (codes === undefined || category === undefined) continue;
    if (category !== "Mn" && category !== "Me" && category !== "Cf") continue;
    const [from, to] = codes.split("..");
    if (from === undefined) continue;
    const start = Number.parseInt(from, 16);
    ranges.push({ start, end: to === undefined ? start : Number.parseInt(to, 16) });
  }
  if (ranges.length === 0) {
    throw new GenerateJoiningTypesError({
      message: "DerivedGeneralCategory.txt yielded no Mn/Me/Cf ranges",
    });
  }
  return ranges;
};

const inRanges = (ranges: Range[], cp: number): boolean =>
  ranges.some((r) => cp >= r.start && cp <= r.end);

/** Merge sorted-by-start ranges, coalescing adjacent and overlapping entries. */
const coalesce = (ranges: Range[]): Range[] => {
  const sorted = [...ranges].sort((a, b) => a.start - b.start);
  const merged: Range[] = [];
  for (const range of sorted) {
    const last = merged.at(-1);
    if (last && range.start <= last.end + 1) {
      last.end = Math.max(last.end, range.end);
      continue;
    }
    merged.push({ ...range });
  }
  return merged;
};

const rangesFromCodePoints = (codePoints: number[]): Range[] =>
  coalesce([...codePoints].sort((a, b) => a - b).map((cp) => ({ start: cp, end: cp })));

type Tables = {
  /** Joining_Type D, L or C: can connect to the FOLLOWING character. */
  forward: Range[];
  /** Joining_Type D, R or C: can connect to the PRECEDING character. */
  backward: Range[];
  /** Joining_Type T: skipped when looking for a joining neighbour. */
  transparent: Range[];
  /** Joining_Type D, L or R: a cursive letter, used to gate the whole path. */
  letters: Range[];
};

const buildTables = (explicit: Map<number, JoiningType>, transparentDefaults: Range[]): Tables => {
  const forward: number[] = [];
  const backward: number[] = [];
  const letters: number[] = [];
  const explicitTransparent: number[] = [];

  for (const [cp, joining] of explicit) {
    if (joining === "D" || joining === "L" || joining === "C") forward.push(cp);
    if (joining === "D" || joining === "R" || joining === "C") backward.push(cp);
    if (joining === "D" || joining === "L" || joining === "R") letters.push(cp);
    if (joining === "T") explicitTransparent.push(cp);
  }

  // An explicit ArabicShaping entry overrides the Mn/Me/Cf default, so drop any
  // default-transparent code point the file classifies as something else.
  const overridden = new Set(
    [...explicit.entries()].filter(([, j]) => j !== "T").map(([cp]) => cp),
  );
  const transparent: Range[] = [];
  for (const range of coalesce([
    ...transparentDefaults,
    ...rangesFromCodePoints(explicitTransparent),
  ])) {
    let start: number | null = null;
    for (let cp = range.start; cp <= range.end; cp++) {
      if (overridden.has(cp)) {
        if (start !== null) {
          transparent.push({ start, end: cp - 1 });
          start = null;
        }
        continue;
      }
      start ??= cp;
    }
    if (start !== null) transparent.push({ start, end: range.end });
  }

  return {
    forward: rangesFromCodePoints(forward),
    backward: rangesFromCodePoints(backward),
    transparent: coalesce(transparent),
    letters: rangesFromCodePoints(letters),
  };
};

/** Flatten ranges to a `[start, end, start, end, …]` literal for binary search. */
const formatRanges = (ranges: Range[]): string => {
  const parts: string[] = [];
  for (const { start, end } of ranges) {
    parts.push(`0x${start.toString(16)}, 0x${end.toString(16)},`);
  }
  const lines: string[] = [];
  for (let i = 0; i < parts.length; i += 6) {
    lines.push(`  ${parts.slice(i, i + 6).join(" ")}`);
  }
  return lines.join("\n");
};

const renderModule = (
  tables: Tables,
  versions: { shaping: string; category: string },
): string => `/**
 * GENERATED FILE — do not edit.
 *
 * Unicode Joining_Type ranges, derived from the pinned UCD sources by
 * \`scripts/generate-joining-types.ts\`. Regenerate with:
 *
 *   bun run generate:joining-types
 *
 * ArabicShaping.txt: ${versions.shaping}
 * DerivedGeneralCategory.txt: ${versions.category}
 *
 * Each table is a flat sorted \`[start, end, …]\` list of inclusive code-point
 * ranges. Transparent is dominated by the General_Category Mn/Me/Cf default,
 * which is why it is by far the largest table.
 */

/** Joining_Type D, L or C: connects to the FOLLOWING character. */
export const JOINS_FORWARD_RANGES: readonly number[] = [
${formatRanges(tables.forward)}
];

/** Joining_Type D, R or C: connects to the PRECEDING character. */
export const JOINS_BACKWARD_RANGES: readonly number[] = [
${formatRanges(tables.backward)}
];

/** Joining_Type T: transparent, skipped when looking for a joining neighbour. */
export const JOINING_TRANSPARENT_RANGES: readonly number[] = [
${formatRanges(tables.transparent)}
];

/** Joining_Type D, L or R: a cursive letter. Gates the whole joining path. */
export const JOINING_LETTER_RANGES: readonly number[] = [
${formatRanges(tables.letters)}
];
`;

const main = async (): Promise<void> => {
  const mode = process.argv[2];
  if (mode !== "write" && mode !== "check") {
    throw new GenerateJoiningTypesError({
      message: "Usage: bun scripts/generate-joining-types.ts <write|check>",
    });
  }

  const sources = await readSourceEntries();
  const shapingSource = requireSource(sources, ARABIC_SHAPING_ID);
  const categorySource = requireSource(sources, GENERAL_CATEGORY_ID);

  const explicit = parseArabicShaping(await readCachedSource(shapingSource));
  const transparentDefaults = parseTransparentDefaults(await readCachedSource(categorySource));

  // Two spot checks against the standard's own tables, so a parser regression
  // fails here rather than surfacing as mis-shaped text much later.
  if (explicit.get(0x06_28) !== "D") {
    throw new GenerateJoiningTypesError({
      message: "ARABIC LETTER BEH (U+0628) must be Dual_Joining",
    });
  }
  if (explicit.get(0x06_27) !== "R") {
    throw new GenerateJoiningTypesError({
      message: "ARABIC LETTER ALEF (U+0627) must be Right_Joining",
    });
  }
  if (!inRanges(transparentDefaults, 0x06_4e)) {
    throw new GenerateJoiningTypesError({
      message: "ARABIC FATHA (U+064E) must default to Transparent via Mn",
    });
  }

  const tables = buildTables(explicit, transparentDefaults);
  const rendered = renderModule(tables, {
    shaping: shapingSource.version,
    category: categorySource.version,
  });

  if (mode === "write") {
    await writeFile(OUTPUT_PATH, rendered, "utf8");
    process.stdout.write(
      `joiningTypes.gen.ts written (forward ${tables.forward.length}, backward ${tables.backward.length}, transparent ${tables.transparent.length}, letters ${tables.letters.length} ranges)\n`,
    );
    return;
  }

  const existing = await readFile(OUTPUT_PATH, "utf8").catch(() => null);
  if (existing === null || digest(existing) !== digest(rendered)) {
    throw new GenerateJoiningTypesError({
      message: "joiningTypes.gen.ts is stale. Run `bun run generate:joining-types`.",
    });
  }
  process.stdout.write("joiningTypes.gen.ts is up to date\n");
};

await main();
