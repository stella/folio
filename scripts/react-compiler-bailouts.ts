// React Compiler bailout guard.
//
// The compiler memoizes compatible components automatically, but Folio's
// imperative ProseMirror bindings make some components ineligible. Manual
// useMemo/useCallback calls remain load-bearing in those files, so guard a
// baseline against removing them or introducing an unacknowledged bailout.
//
// Modes:
//   bun scripts/react-compiler-bailouts.ts                  report bailouts
//   bun scripts/react-compiler-bailouts.ts --write-baseline regenerate baseline
//   bun scripts/react-compiler-bailouts.ts --check          CI gate
import { transformSync } from "@babel/core";
import reactCompiler, { type LoggerEvent } from "babel-plugin-react-compiler";
import { readFileSync, writeFileSync } from "node:fs";

const BASELINE_PATH = "scripts/react-compiler-bailouts.json";
const SOURCE_GLOB = "packages/react/src/**/*.{ts,tsx}";

let mode = "report";
if (process.argv.includes("--write-baseline")) {
  mode = "write";
} else if (process.argv.includes("--check")) {
  mode = "check";
}

const files = Array.from(new Bun.Glob(SOURCE_GLOB).scanSync(".")).filter(
  (file) =>
    !file.includes("/__tests__/") &&
    !file.includes(".test.") &&
    !file.includes(".spec.") &&
    !file.endsWith(".gen.ts") &&
    !file.endsWith(".gen.tsx"),
);

const countManualMemos = (code: string): number =>
  (code.match(/\b(?:useMemo|useCallback)\(/gu) ?? []).length;

type Bailout = { memos: number; reasons: Set<string> };

const bailouts = new Map<string, Bailout>();

for (const file of files) {
  const code = readFileSync(file, "utf8");
  const reasons = new Set<string>();
  const logger = {
    logEvent(_filename: string | null, event: LoggerEvent): void {
      if (event.kind === "CompileSkip") {
        reasons.add(`${event.kind}: ${event.reason}`);
        return;
      }
      if (event.kind === "CompileError" || event.kind === "PipelineError") {
        reasons.add(event.kind);
      }
    },
  };

  try {
    transformSync(code, {
      filename: file,
      babelrc: false,
      configFile: false,
      code: false,
      ast: false,
      parserOpts: { plugins: ["typescript", "jsx"] },
      plugins: [[reactCompiler, { target: "18", panicThreshold: "none", logger }]],
    });
  } catch (error: unknown) {
    const message = error instanceof Error ? error.message : String(error);
    reasons.add(`transform-threw: ${message.slice(0, 80)}`);
  }

  if (reasons.size > 0) {
    bailouts.set(file, { reasons, memos: countManualMemos(code) });
  }
}

const current = Object.fromEntries(
  Array.from(bailouts.entries())
    .toSorted(([left], [right]) => left.localeCompare(right))
    .map(([file, { memos }]) => [file, memos]),
);

if (mode === "write") {
  writeFileSync(BASELINE_PATH, `${JSON.stringify(current, null, 2)}\n`);
  console.log(`Wrote ${Object.keys(current).length} bailout files to ${BASELINE_PATH}`);
  process.exit(0);
}

if (mode === "report") {
  for (const [file, { reasons }] of Array.from(bailouts.entries()).toSorted(([left], [right]) =>
    left.localeCompare(right),
  )) {
    console.log(`${file}\t${Array.from(reasons).join(", ")}`);
  }
  console.log(`\nscanned ${files.length} files; ${bailouts.size} contain compiler bailouts`);
  process.exit(0);
}

const baseline = JSON.parse(readFileSync(BASELINE_PATH, "utf8")) as Record<string, number>;
const failures: string[] = [];

for (const [file, memos] of Object.entries(current)) {
  if (!(file in baseline)) {
    failures.push(`${file}: new bailout file`);
    continue;
  }
  if (memos < baseline[file]!) {
    failures.push(`${file}: manual memoization decreased from ${baseline[file]} to ${memos}`);
  }
}

if (failures.length > 0) {
  console.error("React Compiler bailout baseline changed:");
  for (const failure of failures) {
    console.error(`  ${failure}`);
  }
  console.error(
    "Restore load-bearing memoization or verify the change and regenerate the baseline with --write-baseline.",
  );
  process.exit(1);
}

console.log(
  `React Compiler bailout guard: ${Object.keys(current).length} files, memoization intact.`,
);
