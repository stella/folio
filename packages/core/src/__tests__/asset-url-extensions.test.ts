/**
 * Architecture test — no `new URL(..., import.meta.url)` build asset points at
 * a source extension.
 *
 * Defence in depth alongside the `folio-asset-urls/no-source-extension-url`
 * oxlint rule. Lint can be silenced (config edit, file-scoped disable, or a
 * config typo); this test asserts the same invariant on the actual file tree
 * across every package's shipped source and fails the suite regardless of lint
 * state.
 *
 * A worker or asset referenced via `new URL("<spec>", import.meta.url)` is
 * resolved literally by the consuming bundler against the emitted module. The
 * package build renames `*.ts` -> `*.js` in dist, so a specifier that still
 * points at a source extension resolves to a file that never ships and aborts a
 * downstream Vite/rolldown build with UNRESOLVED_ENTRY (the class of failure
 * `font-metrics.worker.ts` caused). Reference the emitted `.js` instead.
 *
 * Negative tests use synthetic fixture strings to prove the analyser flags the
 * offending construction; the lint plugin's leading comment lists the same
 * cases.
 */

import { Glob } from "bun";
import { describe, expect, test } from "bun:test";
import { readFileSync } from "node:fs";
import path from "node:path";

const PACKAGES_DIR = path.resolve(import.meta.dir, "..", "..", "..");

const SOURCE_EXTENSIONS = [".ts", ".tsx", ".cts", ".mts", ".jsx"];

// Match `new URL("<spec>", import.meta.url)` and capture the string specifier.
// Bounded character classes keep matching linear.
const URL_IMPORT_META_REGEX =
  /new\s+URL\(\s*["'](?<spec>[^"']+)["']\s*,\s*import\.meta\.url\s*\)/gu;

// A specifier that names a build asset with a source extension: not an absolute
// URL (`https:`, `data:`), and its path (minus any `?query` / `#hash`) ends in
// a source extension the build renames.
const isSourceExtensionAsset = (spec: string): boolean => {
  if (/^[a-z][a-z0-9+.-]*:/iu.test(spec) || spec.startsWith("//")) {
    return false;
  }
  const base = spec.split(/[?#]/u)[0] ?? spec;
  return SOURCE_EXTENSIONS.some((ext) => base.endsWith(ext));
};

// Strip block and line comments before the raw-text scan so a comment that
// mentions such a URL (this test and the plugin doc both do) cannot trip it.
// The line-comment pattern keeps the character before `//` so it does not eat a
// `:` from a `https://` URL.
const stripComments = (source: string): string =>
  source.replaceAll(/\/\*[\s\S]*?\*\//gu, "").replaceAll(/(?<lead>^|[^:])\/\/[^\n]*/gmu, "$<lead>");

type AssetUrl = { importer: string; specifier: string };

const sourceExtensionUrlsForSource = (filePath: string, source: string): AssetUrl[] => {
  const found: AssetUrl[] = [];
  for (const match of stripComments(source).matchAll(URL_IMPORT_META_REGEX)) {
    const specifier = match.groups?.["spec"];
    if (specifier !== undefined && isSourceExtensionAsset(specifier)) {
      found.push({ importer: filePath, specifier });
    }
  }
  return found;
};

const isTestFile = (relativePath: string): boolean =>
  relativePath.includes("__tests__/") ||
  relativePath.endsWith(".test.ts") ||
  relativePath.endsWith(".test.tsx");

const collectShippedSourceFiles = (): string[] => {
  const glob = new Glob("*/src/**/*.{ts,tsx}");
  const files: string[] = [];
  for (const relative of glob.scanSync({ cwd: PACKAGES_DIR })) {
    if (isTestFile(relative)) {
      continue;
    }
    files.push(path.resolve(PACKAGES_DIR, relative));
  }
  return files;
};

describe("no source-extension URL build assets", () => {
  test("no shipped source file points a new URL(import.meta.url) asset at a source extension", () => {
    const violations = collectShippedSourceFiles().flatMap((file) =>
      sourceExtensionUrlsForSource(file, readFileSync(file, "utf-8")),
    );
    if (violations.length > 0) {
      const formatted = violations
        .map((v) => `  ${v.importer.replace(PACKAGES_DIR, "<packages>")} -> "${v.specifier}"`)
        .join("\n");
      throw new Error(
        "A `new URL(..., import.meta.url)` build asset points at a source " +
          `extension the package build does not emit:\n${formatted}\n\n` +
          "Reference the emitted name (`.js`); in-repo source builds still " +
          "resolve `.js` -> `.ts` via TypeScript extension resolution.",
      );
    }
    expect(violations).toEqual([]);
  });
});

describe("source-extension URL — synthetic fixtures", () => {
  const importer = path.resolve(PACKAGES_DIR, "core/src/layout-engine/measure/measureWorker.ts");

  test("a worker URL at a .ts target is flagged", () => {
    const violations = sourceExtensionUrlsForSource(
      importer,
      'const w = new Worker(new URL("font-metrics.worker.ts", import.meta.url));',
    );
    expect(violations).toHaveLength(1);
    expect(violations[0]?.specifier).toBe("font-metrics.worker.ts");
  });

  test("a .tsx target is flagged", () => {
    const violations = sourceExtensionUrlsForSource(
      importer,
      'new URL("./thing.tsx", import.meta.url);',
    );
    expect(violations).toHaveLength(1);
  });

  test("the emitted .js target is allowed", () => {
    const violations = sourceExtensionUrlsForSource(
      importer,
      'new Worker(new URL("font-metrics.worker.js", import.meta.url));',
    );
    expect(violations).toEqual([]);
  });

  test("an absolute URL with a .ts path is not a build asset", () => {
    const violations = sourceExtensionUrlsForSource(
      importer,
      'new URL("https://example.com/x.ts");',
    );
    expect(violations).toEqual([]);
  });

  test("a source-extension asset carrying a ?query suffix is still flagged", () => {
    const violations = sourceExtensionUrlsForSource(
      importer,
      'new URL("./worker.ts?worker", import.meta.url);',
    );
    expect(violations).toHaveLength(1);
  });

  test("a comment mentioning a .ts URL is ignored", () => {
    const violations = sourceExtensionUrlsForSource(
      importer,
      '// old: new URL("font-metrics.worker.ts", import.meta.url)\nconst x = 1;',
    );
    expect(violations).toEqual([]);
  });
});
