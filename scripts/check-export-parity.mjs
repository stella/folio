#!/usr/bin/env node
/**
 * Fail if the React and Vue adapters silently drift on either:
 *   1. `package.json` `exports` subpaths
 *   2. Named exports from `src/index.ts`
 *
 * Intentional, documented divergences (framework-native differences and
 * not-yet-ported surfaces) are opted out via first-backtick list items in
 *   scripts/parity/export-divergence.md
 *
 * Adapted from the upstream docx-editor export-parity gate for the folio fork.
 */

import { readFileSync, existsSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { collectNamedExports } from "./lib/named-exports.mjs";
import { diffSets, formatDiff } from "./lib/parity-report.mjs";

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const REACT_PKG = resolve(ROOT, "packages/react/package.json");
const VUE_PKG = resolve(ROOT, "packages/vue/package.json");
const REACT_INDEX = resolve(ROOT, "packages/react/src/index.ts");
const VUE_INDEX = resolve(ROOT, "packages/vue/src/index.ts");
const OPT_OUT = resolve(ROOT, "scripts/parity/export-divergence.md");

function exportSubpaths(pkgPath) {
  return new Set(Object.keys(JSON.parse(readFileSync(pkgPath, "utf8")).exports ?? {}));
}

function loadAllowedDivergences() {
  if (!existsSync(OPT_OUT)) return new Set();
  const allowed = new Set();
  for (const line of readFileSync(OPT_OUT, "utf8").split("\n")) {
    const m = line.match(/^\s*-\s+`([^`]+)`/);
    if (m) allowed.add(m[1]);
  }
  return allowed;
}

const allowed = loadAllowedDivergences();
let failed = false;

// 1) Subpath parity
{
  const reactSubpaths = exportSubpaths(REACT_PKG);
  const vueSubpaths = exportSubpaths(VUE_PKG);
  const { leftOnly, rightOnly } = diffSets(reactSubpaths, vueSubpaths, allowed);

  if (leftOnly.length || rightOnly.length) {
    failed =
      formatDiff({
        label: "subpath parity (package.json `exports`)",
        leftLabel: "React-only",
        rightLabel: "Vue-only",
        leftOnly,
        rightOnly,
        strict: true,
      }) || failed;
  } else {
    const common = [...reactSubpaths].filter((s) => vueSubpaths.has(s)).length;
    console.log(
      `OK subpath parity: ${common} shared subpaths (${allowed.size} documented divergences)`,
    );
  }
}

// 2) Named-export parity
{
  const reactNames = collectNamedExports(REACT_INDEX);
  const vueNames = collectNamedExports(VUE_INDEX);
  const { leftOnly, rightOnly } = diffSets(reactNames, vueNames, allowed);

  if (leftOnly.length || rightOnly.length) {
    failed =
      formatDiff({
        label: "named-export parity (src/index.ts)",
        leftLabel: "react-only",
        rightLabel: "vue-only",
        leftOnly,
        rightOnly,
        strict: true,
      }) || failed;
  } else {
    const common = [...reactNames].filter((n) => vueNames.has(n)).length;
    console.log(`OK named-export parity: ${common} shared names match`);
  }
}

if (failed) {
  console.error(
    "Resolution: add the missing surface to the lagging adapter, or document the\n" +
      "intentional divergence in scripts/parity/export-divergence.md",
  );
  process.exit(1);
}
