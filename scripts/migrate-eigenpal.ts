#!/usr/bin/env bun
import { readdir, readFile, stat, writeFile } from "node:fs/promises";
import path from "node:path";

type Finding = {
  file: string;
  line: number;
  message: string;
};

type TransformResult = {
  changed: boolean;
  text: string;
  findings: Finding[];
};

type Options = {
  write: boolean;
  check: boolean;
  paths: string[];
};

const DEFAULT_PATHS = ["."];

const TEXT_EXTENSIONS = new Set([
  ".cjs",
  ".css",
  ".cts",
  ".js",
  ".json",
  ".jsx",
  ".mjs",
  ".mts",
  ".ts",
  ".tsx",
  ".vue",
]);

const IGNORED_DIRS = new Set([
  ".git",
  ".next",
  ".nuxt",
  ".output",
  "coverage",
  "dist",
  "node_modules",
]);

const IGNORED_FILES = new Set([
  "bun.lock",
  "bun.lockb",
  "npm-shrinkwrap.json",
  "package-lock.json",
  "pnpm-lock.yaml",
  "yarn.lock",
]);

const SAFE_SPECIFIER_REPLACEMENTS = new Map([
  ["@eigenpal/docx-editor-react", "@stll/folio-react/compat/eigenpal"],
  ["@eigenpal/docx-editor-react/styles.css", "@stll/folio-react/standalone.css"],
  ["@eigenpal/docx-editor-vue", "@stll/folio-vue"],
  ["@eigenpal/docx-editor-vue/styles.css", "@stll/folio-vue/editor.css"],
  ["@eigenpal/docx-editor-vue/ui", "@stll/folio-vue/ui"],
  ["@eigenpal/docx-editor-vue/composables", "@stll/folio-vue/composables"],
  ["@eigenpal/docx-editor-vue/dialogs", "@stll/folio-vue/dialogs"],
  ["@eigenpal/docx-editor-vue/styles", "@stll/folio-vue/styles"],
  ["@eigenpal/docx-editor-core", "@stll/folio-core/compat/eigenpal"],
  ["@eigenpal/docx-editor-core/types/document", "@stll/folio-core/types/document"],
  ["@eigenpal/docx-editor-core/types/content", "@stll/folio-core/types/content"],
  ["@eigenpal/docx-editor-core/layout-engine", "@stll/folio-core/layout-engine"],
  ["@eigenpal/docx-editor-core/layout-painter", "@stll/folio-core/layout-painter"],
  ["@eigenpal/docx-editor-core/prosemirror", "@stll/folio-core/prosemirror"],
  ["@eigenpal/docx-editor-core/prosemirror/conversion", "@stll/folio-core/prosemirror/conversion"],
  ["@eigenpal/docx-editor-core/prosemirror/extensions", "@stll/folio-core/prosemirror/extensions"],
  ["@eigenpal/docx-editor-core/prosemirror/schema", "@stll/folio-core/prosemirror/schema"],
  [
    "@eigenpal/docx-editor-core/prosemirror/utils/extractTrackedChanges",
    "@stll/folio-core/prosemirror/utils/extractTrackedChanges",
  ],
  [
    "@eigenpal/docx-editor-core/prosemirror/utils/visualLineNavigation",
    "@stll/folio-core/prosemirror/utils/visualLineNavigation",
  ],
  ["@eigenpal/docx-editor-i18n", "@stll/folio-react/compat/eigenpal"],
  ["@eigenpal/nuxt-docx-editor", "@stll/folio-nuxt"],
]);

const PACKAGE_JSON_REPLACEMENTS = new Map<string, { name: string; version: string }>([
  ["@eigenpal/docx-editor-react", { name: "@stll/folio-react", version: "latest" }],
  ["@eigenpal/docx-editor-vue", { name: "@stll/folio-vue", version: "latest" }],
  ["@eigenpal/nuxt-docx-editor", { name: "@stll/folio-nuxt", version: "latest" }],
  ["@eigenpal/docx-editor-core", { name: "@stll/folio-core", version: "latest" }],
  ["@eigenpal/docx-editor-agents", { name: "@stll/folio-agents", version: "latest" }],
  ["@eigenpal/docx-editor-i18n", { name: "@stll/folio-react", version: "latest" }],
]);

const MANUAL_SPECIFIER_MESSAGES: Array<[RegExp, string]> = [
  [
    /^@eigenpal\/docx-editor-react\/(ui|dialogs|hooks|plugin-api|styles)$/u,
    "React subpath has no direct folio equivalent yet. Import from @stll/folio-react root where available, or add a compatibility export.",
  ],
  [
    /^@eigenpal\/docx-editor-vue\/plugin-api$/u,
    "Vue plugin-api subpath has no direct folio equivalent yet. Port plugin usage or add a compatibility export.",
  ],
  [
    /^@eigenpal\/docx-editor-core(\/.*)?$/u,
    "Core public API is not drop-in. Replace with @stll/folio-core only after checking the imported symbols and subpath.",
  ],
  [
    /^@eigenpal\/docx-editor-agents(\/.*)?$/u,
    "Agent API changed. Port DocxReviewer/agentTools/executeToolCall usage to FolioDocxReviewer, bridges, and executeFolioToolCall.",
  ],
  [
    /^@eigenpal\/docx-editor-i18n\/.+$/u,
    "Locale subpaths cannot be rewritten safely. Import locale constants from @stll/folio-react/compat/eigenpal instead.",
  ],
];

export function transformSource(file: string, input: string): TransformResult {
  if (path.basename(file) === "package.json") {
    return transformPackageJson(file, input);
  }

  const findings: Finding[] = [];
  let changed = false;
  const text = input.replace(
    /(?<quote>["'])(?<specifier>@eigenpal\/(?:docx-editor-(?:react|vue|core|agents|i18n)|nuxt-docx-editor)(?:\/[^"']*)?)(\k<quote>)/gu,
    (match, quote: string, specifier: string) => {
      const replacement = SAFE_SPECIFIER_REPLACEMENTS.get(specifier);
      if (replacement) {
        changed = true;
        return `${quote}${replacement}${quote}`;
      }

      addManualFinding(findings, file, input, specifier);
      return match;
    },
  );

  return { changed, text, findings };
}

function transformPackageJson(file: string, input: string): TransformResult {
  const findings: Finding[] = [];
  let json: unknown;
  try {
    json = JSON.parse(input);
  } catch {
    return { changed: false, text: input, findings };
  }

  if (!isRecord(json)) {
    return { changed: false, text: input, findings };
  }

  let changed = false;
  for (const section of [
    "dependencies",
    "devDependencies",
    "peerDependencies",
    "optionalDependencies",
  ]) {
    const deps = json[section];
    if (!isRecord(deps)) {
      continue;
    }

    let sectionChanged = false;
    const nextDeps: Record<string, unknown> = {};
    for (const [name, version] of Object.entries(deps)) {
      const replacement = PACKAGE_JSON_REPLACEMENTS.get(name);
      if (replacement) {
        if (!(replacement.name in deps) && !(replacement.name in nextDeps)) {
          nextDeps[replacement.name] = replacement.version;
        }
        sectionChanged = true;
        changed = true;
      } else {
        nextDeps[name] = version;
      }
    }

    if (sectionChanged) {
      json[section] = nextDeps;
    }

    if ("@eigenpal/docx-editor-i18n" in nextDeps) {
      findings.push({
        file,
        line: lineOf(input, "@eigenpal/docx-editor-i18n"),
        message:
          "@eigenpal/docx-editor-i18n has no package-name replacement. Use @stll/folio-react/messages or @stll/folio-vue/messages.",
      });
    }
  }

  if (!changed) {
    return { changed: false, text: input, findings };
  }

  return {
    changed: true,
    text: `${JSON.stringify(json, null, 2)}\n`,
    findings,
  };
}

function addManualFinding(findings: Finding[], file: string, input: string, specifier: string) {
  for (const [pattern, message] of MANUAL_SPECIFIER_MESSAGES) {
    if (!pattern.test(specifier)) {
      continue;
    }
    const finding = {
      file,
      line: lineOf(input, specifier),
      message: `${specifier}: ${message}`,
    };
    if (!findings.some((existing) => isSameFinding(existing, finding))) {
      findings.push(finding);
    }
    return;
  }
}

function isSameFinding(left: Finding, right: Finding): boolean {
  return left.file === right.file && left.line === right.line && left.message === right.message;
}

function lineOf(input: string, needle: string): number {
  const index = input.indexOf(needle);
  if (index === -1) {
    return 1;
  }
  return input.slice(0, index).split("\n").length;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return value !== null && typeof value === "object" && !Array.isArray(value);
}

async function collectFiles(inputPath: string): Promise<string[]> {
  const info = await stat(inputPath);
  if (info.isFile()) {
    return shouldScanFile(inputPath) ? [inputPath] : [];
  }
  if (!info.isDirectory()) {
    return [];
  }

  const files: string[] = [];
  for (const entry of await readdir(inputPath, { withFileTypes: true })) {
    if (entry.isDirectory() && IGNORED_DIRS.has(entry.name)) {
      continue;
    }
    const fullPath = path.join(inputPath, entry.name);
    if (entry.isDirectory()) {
      files.push(...(await collectFiles(fullPath)));
    } else if (entry.isFile() && shouldScanFile(fullPath)) {
      files.push(fullPath);
    }
  }
  return files;
}

function shouldScanFile(file: string): boolean {
  if (IGNORED_FILES.has(path.basename(file))) {
    return false;
  }

  return path.basename(file) === "package.json" || TEXT_EXTENSIONS.has(path.extname(file));
}

function parseArgs(argv: string[]): Options {
  const paths: string[] = [];
  let write = false;
  let check = false;

  for (const arg of argv) {
    if (arg === "--write") {
      write = true;
    } else if (arg === "--check") {
      check = true;
    } else if (arg === "--help" || arg === "-h") {
      printHelp();
      process.exit(0);
    } else {
      paths.push(arg);
    }
  }

  return { write, check, paths: paths.length > 0 ? paths : DEFAULT_PATHS };
}

function printHelp() {
  console.log(`Usage: bun scripts/migrate-eigenpal.ts [--write] [--check] [paths...]

Rewrites safe Eigenpal docx-editor imports to folio imports and reports manual
migration gaps. Defaults to dry-run mode.

Examples:
  bun run codemod:eigenpal -- apps/web
  bun run codemod:eigenpal -- --write src package.json
  bun run codemod:eigenpal -- --check .
`);
}

async function main() {
  const options = parseArgs(process.argv.slice(2));
  const files = (
    await Promise.all(options.paths.map((inputPath) => collectFiles(path.resolve(inputPath))))
  )
    .flat()
    .sort();

  let changedCount = 0;
  const findings: Finding[] = [];

  for (const file of files) {
    const source = await readFile(file, "utf8");
    if (!source.includes("@eigenpal/")) {
      continue;
    }

    const result = transformSource(file, source);
    findings.push(...result.findings);
    if (!result.changed) {
      continue;
    }

    changedCount += 1;
    if (options.write) {
      await writeFile(file, result.text);
    }
    console.log(
      `${options.write ? "updated" : "would update"} ${path.relative(process.cwd(), file)}`,
    );
  }

  for (const finding of findings) {
    console.warn(
      `${path.relative(process.cwd(), finding.file)}:${finding.line}: ${finding.message}`,
    );
  }

  if (changedCount === 0 && findings.length === 0) {
    console.log("No Eigenpal docx-editor imports found.");
  } else {
    console.log(
      `${options.write ? "Updated" : "Would update"} ${changedCount} file(s); ${findings.length} manual item(s).`,
    );
  }

  if (options.check && (changedCount > 0 || findings.length > 0)) {
    process.exit(1);
  }
}

if (import.meta.main) {
  await main();
}
