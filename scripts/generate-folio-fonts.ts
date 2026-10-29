/**
 * Generate `@stll/folio-fonts` from the pinned `@fontsource` packages.
 *
 * The package ships exactly the faces folio uses, so a consumer installs
 * folio's packages and never an `@fontsource` package itself. Which families,
 * weights, styles and measured subsets those are comes from one table,
 * `packages/core/src/fonts/bundledFonts.ts`; the `@fontsource` packages are
 * devDependencies of `packages/fonts` and are read only here.
 *
 * Written under `packages/fonts`:
 *
 *   fonts/<pkg>/files/*.woff2     every subset each served face is cut into,
 *                                 which is what a browser loads
 *   fonts/<pkg>/files/*.woff      the measured subsets only, which is what the
 *                                 headless measurer and the PDF writer parse
 *   fonts/<pkg>/unicode.json      the ranges each subset was cut with
 *   fonts/<pkg>/LICENSE           the family's licence, beside its files
 *   licenses/<pkg>.txt            the same licence, collected in one place
 *   fonts.css                     the `@font-face` rules `@fontsource`'s own
 *                                 stylesheets declare, with urls relative to
 *                                 the package: `.woff2`, plus the `.woff`
 *                                 fallback for the measured subsets only
 *   THIRD_PARTY_NOTICES.md        each family, its source, licence and role
 *   src/catalogue.gen.ts          the same table for the package's own code
 *
 * Usage:
 *   bun scripts/generate-folio-fonts.ts write
 *   bun scripts/generate-folio-fonts.ts check
 */

import { TaggedError } from "better-result";
import { existsSync } from "node:fs";
import { mkdir, readdir, readFile, rm, writeFile } from "node:fs/promises";
import { createRequire } from "node:module";
import path from "node:path";

import {
  BUNDLED_FAMILY_SUBSTITUTES,
  BUNDLED_STYLES,
  BUNDLED_WEIGHTS,
  type BundledFamily,
  bundledFontFilePath,
  FONTSOURCE_PACKAGES,
  MEASURED_SUBSETS,
  UNICODE_RANGE_FILE,
} from "../packages/core/src/fonts/bundledFonts";

const REPO_ROOT = path.resolve(import.meta.dir, "..");
const PACKAGE_ROOT = path.join(REPO_ROOT, "packages", "fonts");
const SCRIPT = "generate:folio-fonts";

/** Directories the generator owns outright: anything else in them is stale. */
const OWNED_DIRECTORIES = ["fonts", "licenses"] as const;

const STYLESHEET = "fonts.css";
const NOTICES = "THIRD_PARTY_NOTICES.md";
const CATALOGUE = "src/catalogue.gen.ts";

/**
 * Where each family is developed. `@fontsource` repackages the files Google
 * Fonts serves; the notices name the project the faces come from.
 */
const SOURCE_PROJECTS: Readonly<Record<BundledFamily, string>> = {
  Arimo: "https://github.com/googlefonts/arimo",
  Caladea: "https://github.com/huertatipografica/Caladea",
  Carlito: "https://github.com/googlefonts/carlito",
  Cousine: "https://github.com/googlefonts/cousine",
  Lato: "https://www.latofonts.com",
  "Noto Sans Arabic": "https://github.com/notofonts/arabic",
  "Source Sans 3": "https://github.com/adobe-fonts/source-sans",
  Tinos: "https://github.com/googlefonts/tinos",
};

const KNOWN_LICENSES = new Set(["OFL-1.1", "Apache-2.0"]);

class GenerateFolioFontsError extends TaggedError("GenerateFolioFontsError")<{
  message: string;
}> {}

const fail = (message: string): never => {
  throw new GenerateFolioFontsError({ message });
};

type Face = {
  family: BundledFamily;
  packageName: string;
  subset: string;
  weight: number;
  style: "normal" | "italic";
  unicodeRange: string;
  woff2: string;
  woff: string | null;
};

type Family = {
  family: BundledFamily;
  packageName: string;
  version: string;
  license: string;
  attribution: string;
};

const isBundledFamily = (name: string): name is BundledFamily =>
  Object.hasOwn(FONTSOURCE_PACKAGES, name);

const MEASURED: ReadonlySet<string> = new Set(MEASURED_SUBSETS);

/** The two fields of an `@fontsource` manifest the notices record. */
const readManifest = async (
  file: string,
): Promise<{ version: string; license: string }> => {
  const manifest: unknown = JSON.parse(await readFile(file, "utf8"));
  if (
    typeof manifest === "object" &&
    manifest !== null &&
    "version" in manifest &&
    typeof manifest.version === "string" &&
    "license" in manifest &&
    typeof manifest.license === "string"
  ) {
    return { version: manifest.version, license: manifest.license };
  }
  return fail(`${file} has no version or license`);
};

const requireFromFonts = createRequire(path.join(PACKAGE_ROOT, "package.json"));

const fontsourceDirectory = (packageName: string): string =>
  path.dirname(
    requireFromFonts.resolve(`@fontsource/${packageName}/package.json`),
  );

/** `400.css`, `400-italic.css`: the stylesheet `@fontsource` ships per variant. */
const variantStylesheet = (weight: number, style: string): string =>
  `${String(weight)}${style === "italic" ? "-italic" : ""}.css`;

const RULE_RE =
  /\/\*\s*(?<name>[a-z0-9-]+)\s*\*\/\s*@font-face\s*\{(?<body>[^}]*)\}/gu;
const DECLARATION_RE = /^\s*(?<property>[a-z-]+)\s*:\s*(?<value>.+?);\s*$/gmu;
const WOFF2_SOURCE_RE =
  /url\(\.\/files\/(?<file>[a-z0-9-]+\.woff2)\)\s*format\('woff2'\)/u;

type Rule = {
  name: string;
  declarations: readonly (readonly [string, string])[];
  woff2File: string;
  unicodeRange: string;
};

const parseRules = (css: string, source: string): readonly Rule[] =>
  [...css.matchAll(RULE_RE)].map((match) => {
    const name = match.groups?.["name"] ?? fail(`${source}: unnamed rule`);
    const body = match.groups?.["body"] ?? "";
    const declarations = [...body.matchAll(DECLARATION_RE)].map(
      (declaration) =>
        [
          declaration.groups?.["property"] ?? "",
          declaration.groups?.["value"] ?? "",
        ] as const,
    );
    const src =
      declarations.find(([property]) => property === "src")?.[1] ??
      fail(`${source}: ${name} has no src`);
    const woff2File =
      WOFF2_SOURCE_RE.exec(src)?.groups?.["file"] ??
      fail(`${source}: ${name} lists no .woff2 file`);
    const unicodeRange =
      declarations.find(([property]) => property === "unicode-range")?.[1] ??
      fail(`${source}: ${name} has no unicode-range`);
    return { name, declarations, woff2File, unicodeRange };
  });

/**
 * A rule as `@fontsource` writes it, its `src` pointed at the package's own
 * files: the `.woff2`, then the `.woff` for a measured subset. A browser takes
 * the first format it reads, so the fallback costs a page nothing, and it is
 * the file a measurement harness parses beside the canvas: the bytes the
 * headless measurer reads.
 */
const renderRule = (
  rule: Rule,
  packageName: string,
  woffFile: string | null,
): string =>
  [
    `/* ${rule.name} */`,
    "@font-face {",
    ...rule.declarations.map(([property, value]) =>
      property === "src"
        ? `  src: url(./fonts/${packageName}/files/${rule.woff2File}) format('woff2')${woffFile === null ? "" : `, url(./fonts/${packageName}/${woffFile}) format('woff')`};`
        : `  ${property}: ${value};`,
    ),
    "}",
  ].join("\n");

const firstLine = (text: string): string =>
  text.split("\n", 1)[0]?.trim() ?? "";

type Output = Map<string, Uint8Array>;

const textBytes = (text: string): Uint8Array => new TextEncoder().encode(text);

const generate = async () => {
  const output: Output = new Map();
  const families: Family[] = [];
  const faces: Face[] = [];
  const stylesheetSections: string[] = [];

  for (const [family, packageName] of Object.entries(FONTSOURCE_PACKAGES)) {
    if (!isBundledFamily(family)) continue;
    const source = fontsourceDirectory(packageName);
    const manifest = await readManifest(path.join(source, "package.json"));
    if (!KNOWN_LICENSES.has(manifest.license)) {
      fail(
        `@fontsource/${packageName} is licensed ${manifest.license}; the package declares OFL-1.1 AND Apache-2.0 only.`,
      );
    }
    const license = await readFile(path.join(source, "LICENSE"));
    families.push({
      family,
      packageName,
      version: manifest.version,
      license: manifest.license,
      attribution: firstLine(new TextDecoder().decode(license)),
    });
    output.set(`fonts/${packageName}/LICENSE`, license);
    output.set(`licenses/${packageName}.txt`, license);
    output.set(
      `fonts/${packageName}/${UNICODE_RANGE_FILE}`,
      await readFile(path.join(source, UNICODE_RANGE_FILE)),
    );

    const rules: string[] = [];
    for (const weight of BUNDLED_WEIGHTS) {
      for (const style of BUNDLED_STYLES) {
        const stylesheet = path.join(source, variantStylesheet(weight, style));
        if (!existsSync(stylesheet)) {
          if (style === "normal") {
            fail(`@fontsource/${packageName} has no ${String(weight)} face`);
          }
          continue;
        }
        const parsed = parseRules(
          await readFile(stylesheet, "utf8"),
          path.relative(REPO_ROOT, stylesheet),
        );
        if (parsed.length === 0) {
          fail(`${stylesheet} declares no @font-face rules`);
        }
        const suffix = `-${String(weight)}-${style}`;
        for (const rule of parsed) {
          const prefix = `${packageName}-`;
          if (!rule.name.startsWith(prefix) || !rule.name.endsWith(suffix)) {
            fail(`${stylesheet}: unexpected rule name ${rule.name}`);
          }
          const subset = rule.name.slice(prefix.length, -suffix.length);
          const woff2 = `files/${rule.woff2File}`;
          output.set(
            `fonts/${packageName}/${woff2}`,
            await readFile(path.join(source, woff2)),
          );
          const woffPath = bundledFontFilePath({
            packageName,
            subset,
            weight,
            style,
            format: "woff",
          });
          const measured = MEASURED.has(subset);
          if (measured) {
            output.set(
              `fonts/${packageName}/${woffPath}`,
              await readFile(path.join(source, woffPath)),
            );
          }
          faces.push({
            family,
            packageName,
            subset,
            weight,
            style,
            unicodeRange: rule.unicodeRange,
            woff2: `fonts/${packageName}/${woff2}`,
            woff: measured ? `fonts/${packageName}/${woffPath}` : null,
          });
          rules.push(renderRule(rule, packageName, measured ? woffPath : null));
        }
      }
    }
    stylesheetSections.push(
      [
        `/* ${family}: stands in for ${BUNDLED_FAMILY_SUBSTITUTES[family]}. ${manifest.license}, see licenses/${packageName}.txt. */`,
        ...rules,
      ].join("\n\n"),
    );
  }

  output.set(
    STYLESHEET,
    textBytes(
      [
        "/*",
        " * @stll/folio-fonts: the document fonts folio measures and paints with.",
        " * Generated by `bun run generate:folio-fonts`; do not edit.",
        " *",
        " * Every face is woff2, split into the subsets and unicode-ranges its",
        " * @fontsource package cuts it into, so a browser downloads a file only",
        " * when a page paints one of its code points. The latin, latin-ext and",
        " * arabic subsets list their .woff after it: the file folio measures",
        " * with headlessly, which a browser never needs. The font files are",
        " * licensed as THIRD_PARTY_NOTICES.md lists; this stylesheet is",
        " * Apache-2.0.",
        " */",
        "",
        stylesheetSections.join("\n\n"),
        "",
      ].join("\n"),
    ),
  );
  output.set(NOTICES, textBytes(renderNotices(families)));
  output.set(CATALOGUE, textBytes(renderCatalogue(families, faces)));
  return { output, faces };
};

const renderNotices = (families: readonly Family[]): string =>
  [
    "# Third-party notices",
    "",
    "`@stll/folio-fonts` redistributes the font files below, unmodified, as the",
    "[Fontsource](https://fontsource.org) packages publish them. Each family keeps",
    "its own licence: the full text is in `licenses/<directory>.txt` and again",
    "beside the files in `fonts/<directory>/LICENSE`. The package's own code and",
    "the generated `fonts.css` are licensed under Apache-2.0 (see `LICENSE`).",
    "",
    "Generated by `bun run generate:folio-fonts`; do not edit.",
    "",
    "| Family | Directory | Source | Licence | Stands in for |",
    "| --- | --- | --- | --- | --- |",
    ...families.map(
      ({ family, packageName, license }) =>
        `| ${family} | \`fonts/${packageName}\` | ${SOURCE_PROJECTS[family]} | ${license} | ${BUNDLED_FAMILY_SUBSTITUTES[family]} |`,
    ),
    "",
    "## Attribution",
    "",
    ...families.flatMap(({ family, packageName, version, attribution }) => [
      `### ${family}`,
      "",
      `From \`@fontsource/${packageName}\` ${version}. ${attribution}`,
      "",
    ]),
  ].join("\n");

const quote = (value: string): string => JSON.stringify(value);

const renderCatalogue = (
  families: readonly Family[],
  faces: readonly Face[],
): string =>
  [
    "// Generated by `bun run generate:folio-fonts` from",
    "// packages/core/src/fonts/bundledFonts.ts and the pinned @fontsource",
    "// packages; do not edit.",
    "",
    'import type { FolioFontFace, FolioFontFamily } from "./types";',
    "",
    "/** The bundled family names. */",
    `export type FolioFontFamilyName =\n${families.map(({ family }) => `  | ${quote(family)}`).join("\n")};`,
    "",
    "/** Each family and the directory under `fonts/` that holds its files. */",
    "export const PACKAGES: Readonly<Record<FolioFontFamilyName, string>> = {",
    ...families.map(
      ({ family, packageName }) => `  ${quote(family)}: ${quote(packageName)},`,
    ),
    "};",
    "",
    "export const FAMILIES: readonly FolioFontFamily[] = [",
    ...families.map(
      ({ family, packageName, version, license }) =>
        `  {\n    family: ${quote(family)},\n    packageName: ${quote(packageName)},\n    license: ${quote(license)},\n    fontsourceVersion: ${quote(version)},\n    substitutes: ${quote(BUNDLED_FAMILY_SUBSTITUTES[family])},\n    source: ${quote(SOURCE_PROJECTS[family])},\n  },`,
    ),
    "];",
    "",
    "export const FACES: readonly FolioFontFace[] = [",
    ...faces.map(
      (face) =>
        `  {\n    family: ${quote(face.family)},\n    packageName: ${quote(face.packageName)},\n    subset: ${quote(face.subset)},\n    weight: ${String(face.weight)},\n    style: ${quote(face.style)},\n    unicodeRange: ${quote(face.unicodeRange)},\n    woff2: ${quote(face.woff2)},\n    woff: ${face.woff === null ? "null" : quote(face.woff)},\n  },`,
    ),
    "];",
    "",
  ].join("\n");

/** Every file under the owned directories, package-relative with `/` separators. */
const listOwnedFiles = async (): Promise<string[]> => {
  const files: string[] = [];
  for (const directory of OWNED_DIRECTORIES) {
    const root = path.join(PACKAGE_ROOT, directory);
    if (!existsSync(root)) continue;
    for (const entry of await readdir(root, {
      recursive: true,
      withFileTypes: true,
    })) {
      if (!entry.isFile()) continue;
      files.push(
        path
          .relative(PACKAGE_ROOT, path.join(entry.parentPath, entry.name))
          .split(path.sep)
          .join("/"),
      );
    }
  }
  return files;
};

const sameBytes = (a: Uint8Array, b: Uint8Array): boolean =>
  a.byteLength === b.byteLength && a.every((byte, index) => byte === b[index]);

const summarize = (output: Output, faces: readonly Face[]) => {
  let woff2 = 0;
  let woff = 0;
  for (const [file, bytes] of output) {
    if (file.endsWith(".woff2")) woff2 += bytes.byteLength;
    else if (file.endsWith(".woff")) woff += bytes.byteLength;
  }
  return `${String(faces.length)} faces, ${String(woff2)} bytes of .woff2, ${String(woff)} bytes of .woff`;
};

const main = async (mode: string | undefined) => {
  if (mode !== "write" && mode !== "check") {
    fail(
      `Usage: bun scripts/generate-folio-fonts.ts [check|write], not "${String(mode)}".`,
    );
  }
  const { output, faces } = await generate();

  if (mode === "write") {
    for (const directory of OWNED_DIRECTORIES) {
      await rm(path.join(PACKAGE_ROOT, directory), {
        recursive: true,
        force: true,
      });
    }
    for (const [file, bytes] of output) {
      const target = path.join(PACKAGE_ROOT, file);
      await mkdir(path.dirname(target), { recursive: true });
      await writeFile(target, bytes);
    }
    process.stdout.write(
      `@stll/folio-fonts written (${summarize(output, faces)})\n`,
    );
    return;
  }

  const problems: string[] = [];
  for (const [file, bytes] of output) {
    const target = path.join(PACKAGE_ROOT, file);
    if (!existsSync(target)) {
      problems.push(`missing ${file}`);
      continue;
    }
    if (!sameBytes(new Uint8Array(await readFile(target)), bytes)) {
      problems.push(`stale ${file}`);
    }
  }
  for (const file of await listOwnedFiles()) {
    if (!output.has(file)) problems.push(`unexpected ${file}`);
  }
  if (problems.length > 0) {
    fail(
      `packages/fonts is out of date with its sources. Run \`bun run ${SCRIPT}\`.\n${problems.map((problem) => `  ${problem}`).join("\n")}`,
    );
  }
  process.stdout.write(
    `@stll/folio-fonts is up to date (${summarize(output, faces)})\n`,
  );
};

await main(process.argv[2]);
