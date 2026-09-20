/**
 * Architecture test — one underline → CSS table.
 *
 * There were four readers of `ST_Underline` and only three of them read a
 * total table: the ProseMirror mark's `toDOM` carried a private four-entry one,
 * so the editor drew a plain line for every member it had never heard of. A
 * rendering decision per enumeration member is a table, and a table has one
 * owner; a second copy is how the editor and the page drift apart again.
 *
 * The scan flags a run of adjacent entries that maps `ST_Underline` members to
 * CSS `text-decoration-style` keywords. A run that would also be a legal
 * `border-style` table is left alone: `ST_Border` spells `single`, `thick`,
 * `double` and `dotted` the same way and owns its own table in
 * `utils/borderCss.ts`.
 */

import { Glob } from "bun";
import { describe, expect, test } from "bun:test";
import { readFileSync } from "node:fs";
import path from "node:path";

import { CSS_BORDER_STYLE_VALUES, CSS_BORDER_STYLES } from "../utils/borderCss";
import { UNDERLINE_DECORATION_CSS } from "../utils/formatToStyle";

const { relative, resolve } = path;

const PACKAGES_DIR = resolve(import.meta.dir, "../../../");

/**
 * The files allowed to hold such a table, each the owner of one vocabulary.
 * `strokes.ts` maps the same members to display-list patterns, which happen to
 * be spelled like the CSS keywords; `outlineDash.property.test.ts` holds that
 * table to `ST_Underline`.
 */
const TABLE_OWNERS = ["core/src/utils/formatToStyle.ts", "core/src/display-list/build/strokes.ts"];

/** The keywords and the members both come from the owner, never a hand list. */
const CSS_DECORATION_STYLES: ReadonlySet<string> = new Set(
  Object.values(UNDERLINE_DECORATION_CSS).flatMap((decoration) =>
    "decorationStyle" in decoration ? [decoration.decorationStyle] : [],
  ),
);
const BORDER_STYLES: ReadonlySet<string> = new Set(Object.keys(CSS_BORDER_STYLES));
const BORDER_KEYWORDS: ReadonlySet<string> = new Set(CSS_BORDER_STYLE_VALUES);

const isUnderlineStyle = (value: string): boolean => Object.hasOwn(UNDERLINE_DECORATION_CSS, value);

/** `dotted: "dotted",` and `dotted: { decorationStyle: "dotted" },` alike. */
const ENTRY_PATTERN = /^\s*(?<key>\w+)\s*:\s*[^"']*["'](?<value>[a-z-]+)["']/u;

type TableEntry = { key: string; value: string };

/** A run of adjacent entries that maps underline members to CSS keywords. */
const underlineCssRuns = (source: string): TableEntry[][] => {
  const runs: TableEntry[][] = [];
  let run: TableEntry[] = [];
  for (const line of source.split("\n")) {
    const { key, value } = ENTRY_PATTERN.exec(line)?.groups ?? {};
    const isTableEntry =
      key !== undefined &&
      value !== undefined &&
      isUnderlineStyle(key) &&
      CSS_DECORATION_STYLES.has(value);
    if (isTableEntry) {
      run.push({ key, value });
      continue;
    }
    if (run.length > 1) {
      runs.push(run);
    }
    run = [];
  }
  if (run.length > 1) {
    runs.push(run);
  }
  // A run that is also a legal `border-style` table belongs to that vocabulary.
  return runs.filter(
    (entries) =>
      !entries.every(({ key, value }) => BORDER_STYLES.has(key) && BORDER_KEYWORDS.has(value)),
  );
};

const isTestFile = (relativePath: string): boolean =>
  relativePath.includes("__tests__/") ||
  relativePath.endsWith(".test.ts") ||
  relativePath.endsWith(".test.tsx");

const scannedFiles = (): string[] => {
  const glob = new Glob("*/src/**/*.{ts,tsx,vue}");
  const files: string[] = [];
  for (const relativePath of glob.scanSync({ cwd: PACKAGES_DIR })) {
    if (isTestFile(relativePath) || TABLE_OWNERS.includes(relativePath)) {
      continue;
    }
    files.push(resolve(PACKAGES_DIR, relativePath));
  }
  return files;
};

describe("one underline → CSS table", () => {
  test("the scan reaches every package's source", () => {
    expect(scannedFiles().length).toBeGreaterThan(1000);
  });

  test("no second table exists outside the owner", () => {
    const violations = scannedFiles().flatMap((file) => {
      const runs = underlineCssRuns(readFileSync(file, "utf-8"));
      return runs.map((entries) => {
        const mapped = entries.map(({ key, value }) => `${key} -> ${value}`).join(", ");
        return `  ${relative(PACKAGES_DIR, file)}: ${mapped}`;
      });
    });
    if (violations.length > 0) {
      throw new Error(
        `An underline style renders through one table, ${TABLE_OWNERS[0]}, but found:\n` +
          `${violations.join("\n")}\n\n` +
          "Call `underlineDecorationCss` instead of deciding again here.",
      );
    }
    expect(violations).toEqual([]);
  });

  test("the scan flags the table this removed", () => {
    const runs = underlineCssRuns(
      [
        "const styleMap: Record<string, string> = {",
        '  double: "double",',
        '  dotted: "dotted",',
        '  dash: "dashed",',
        '  wave: "wavy",',
        "};",
      ].join("\n"),
    );

    expect(runs).toHaveLength(1);
    expect(runs[0]).toHaveLength(4);
  });

  test("the border vocabulary's own table is not flagged", () => {
    const runs = underlineCssRuns(
      [
        '  single: "solid",',
        '  thick: "solid",',
        '  double: "double",',
        '  dotted: "dotted",',
      ].join("\n"),
    );

    expect(runs).toEqual([]);
  });

  test("a single decision is not a table", () => {
    expect(underlineCssRuns('  element.style.textDecorationStyle = "dotted";')).toEqual([]);
    expect(underlineCssRuns('  dotted: "dotted",')).toEqual([]);
  });
});
