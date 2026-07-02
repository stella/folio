// Completeness guard for the self-sufficient `standalone.css`.
//
// `standalone.css` ships a PRE-COMPILED copy of every Tailwind utility folio's
// own React components use, scoped under `.folio-root` (see
// `src/styles/standalone.css` + `../tailwind.config.js`). The compile scans the
// package source; if that scan ever drifts from what actually ships in the dist
// JS (a new component dir the `@source` glob misses, a class assembled in a way
// the scanner cannot see, a stale build), a consumer on the standalone path
// gets an unstyled control with no error.
//
// This pure helper closes that gap: it extracts Tailwind-ish class tokens from
// the packed dist JS (only from `className:` values and `cn()`/`clsx()`/`cx()`
// arguments, so import specifiers and i18n keys are never mistaken for classes)
// and asserts each one has a generated rule in the packed `standalone.css`.
// Kept separate from `validate-dist.ts` so the extraction rules are unit tested
// in `standalone-css-coverage.test.ts`, mirroring `dist-url-targets.ts`.

import type { DistJsFile } from "./dist-url-targets";

export type StandaloneCoverageResult = {
  /** Distinct Tailwind-ish class tokens found in the dist JS. */
  candidates: number;
  /** Class tokens used in the JS that have no rule in `standalone.css`. */
  uncovered: string[];
};

type CoverageInput = {
  jsFiles: DistJsFile[];
  standaloneCss: string;
};

// folio's own semantic class prefixes (styled by `editor.css`, not utilities).
const CUSTOM_CLASS_PREFIXES = [
  "folio",
  "docx",
  "hf-",
  "pg-",
  "paged-editor",
  "prosemirror",
  "image-",
  "layout-",
  "band-",
  "cm-",
  "tiptap",
];

// Marker/inert classes that neither the own-Tailwind path nor the standalone
// path generates a rule for: Tailwind state markers (`group`, `peer`) and
// animation utilities that require a plugin folio does not enable (so they are
// equally no-ops on both paths).
const IGNORED_CLASSES = new Set(["group", "peer", "sr-only", "not-prose", "prose", "dark"]);
const INERT_CLASS_RE =
  /^(?:animate-|slide-(?:in|out)|fade-(?:in|out)|zoom-(?:in|out)|spin$|pulse$)/u;

// Structural shape of a Tailwind utility, optionally prefixed with variants
// (`hover:`, `dark:`, `data-[state=open]:`, `[&_svg]:`) and an opacity/arbitrary
// suffix. Deliberately loose on the utility body; the authoritative check is
// membership in the generated set, this only screens out non-class strings.
const TAILWIND_TOKEN_RE =
  /^(?:(?:[-a-z0-9]*(?:\[[^\]]*\])?[-a-z0-9]*):)*!?-?[a-z][a-z0-9]*(?:-[a-z0-9.]+)*(?:-?\[[^\]]*\])?(?:\/[0-9.]+)?$/u;

// Class selector token in the compiled CSS, capturing the escaped class name.
const SCOPED_UTILITY_RE = /\.folio-root \.((?:\\.|[^\\{,:>~+\s[])+)/gu;
const ANY_CLASS_RE = /\.((?:\\.|[^\s.,{}()>~+:#[\]])+)/gu;

const unescapeClass = (escaped: string): string => escaped.replace(/\\(.)/gu, "$1");

// Pull the contents of every `"..."`, `'...'`, and `` `...` `` literal from a
// snippet; template interpolations (`${...}`) are runtime values, not classes,
// so they are blanked to a separator.
const collectLiteralContents = (snippet: string, out: string[]): void => {
  const re = /(["'`])((?:\\.|(?!\1)[^\\])*)\1/gu;
  for (const match of snippet.matchAll(re)) {
    const quote = match[1];
    const value = match[2] ?? "";
    out.push(quote === "`" ? value.replace(/\$\{[^}]*\}/gu, " ") : value);
  }
};

// Return the argument list of each `cn(` / `clsx(` / `cx(` call, using balanced
// paren matching (string-aware) so nested calls and following JSX are excluded.
const collectCallArgumentBlocks = (code: string): string[] => {
  const blocks: string[] = [];
  const callRe = /\b(?:cn|clsx|cx|twMerge|twJoin)\(/gu;
  let match: RegExpExecArray | null = callRe.exec(code);
  while (match !== null) {
    let i = callRe.lastIndex;
    const start = i;
    let depth = 1;
    while (i < code.length && depth > 0) {
      const c = code[i];
      if (c === '"' || c === "'" || c === "`") {
        i += 1;
        while (i < code.length && code[i] !== c) {
          if (code[i] === "\\") {
            i += 1;
          }
          i += 1;
        }
      } else if (c === "(") {
        depth += 1;
      } else if (c === ")") {
        depth -= 1;
      }
      i += 1;
    }
    blocks.push(code.slice(start, i - 1));
    callRe.lastIndex = i;
    match = callRe.exec(code);
  }
  return blocks;
};

const extractClassStrings = (jsFiles: DistJsFile[]): string[] => {
  const strings: string[] = [];
  const classNameLiteralRe =
    /className:\s*(`(?:\\.|[^`\\])*`|"(?:\\.|[^"\\])*"|'(?:\\.|[^'\\])*')/gu;
  for (const { code } of jsFiles) {
    for (const match of code.matchAll(classNameLiteralRe)) {
      collectLiteralContents(match[1] ?? "", strings);
    }
    for (const block of collectCallArgumentBlocks(code)) {
      collectLiteralContents(block, strings);
    }
  }
  return strings;
};

export const findUncoveredUtilities = ({
  jsFiles,
  standaloneCss,
}: CoverageInput): StandaloneCoverageResult => {
  const generated = new Set<string>();
  for (const match of standaloneCss.matchAll(SCOPED_UTILITY_RE)) {
    generated.add(unescapeClass(match[1] ?? ""));
  }
  const anyCssClass = new Set<string>();
  for (const match of standaloneCss.matchAll(ANY_CLASS_RE)) {
    anyCssClass.add(unescapeClass(match[1] ?? ""));
  }

  const candidates = new Set<string>();
  for (const classString of extractClassStrings(jsFiles)) {
    for (const token of classString.split(/\s+/u)) {
      if (!token || IGNORED_CLASSES.has(token) || INERT_CLASS_RE.test(token)) {
        continue;
      }
      if (CUSTOM_CLASS_PREFIXES.some((prefix) => token.startsWith(prefix))) {
        continue;
      }
      if (!TAILWIND_TOKEN_RE.test(token)) {
        continue;
      }
      // A token that has a non-utility rule in the bundled `editor.css` portion
      // (e.g. a hand-written `.docx-*` doc class reached via `cn`) is covered by
      // that rule, not a utility.
      if (anyCssClass.has(token) && !generated.has(token)) {
        continue;
      }
      candidates.add(token);
    }
  }

  const uncovered = [...candidates].filter((token) => !generated.has(token)).sort();
  return { candidates: candidates.size, uncovered };
};
