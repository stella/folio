#!/usr/bin/env node
// Cross-adapter parity check between @stll/folio-react and @stll/folio-vue.
//
// Reads the `DocxEditorProps`, `DocxEditorRef`, and nested `PagedEditorRef`
// member names directly from each adapter's TypeScript source and applies
// `scripts/parity/parity.contract.json`.
// (Upstream reads committed API-Extractor snapshots; the folio fork has no Vue
// api-report yet, so we parse the source type declarations instead.)
//
// Source of truth:
//   - React: packages/react/src/components/DocxEditor.props.ts
//   - Vue:   packages/vue/src/components/DocxEditor/types.ts
//
// Fails non-zero on any drift the contract does not acknowledge:
//   - A prop/member exists on an adapter but the contract did not classify it.
//   - A `paired` entry is missing from one adapter's type surface.
//   - A `deferredInVue` entry is missing from React (the contract is stale).
//   - A member is classified in more than one bucket.
//
// The contract is the source of truth: adding a prop/method to either adapter
// without updating the contract is the failure mode this gate exists for.

import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const repoRoot = path.resolve(__dirname, "..");

const REACT_PROPS_SRC = path.join(repoRoot, "packages/react/src/components/DocxEditor.props.ts");
const VUE_TYPES_SRC = path.join(repoRoot, "packages/vue/src/components/DocxEditor/types.ts");
const REACT_PAGED_REF_SRC = path.join(repoRoot, "packages/react/src/paged-editor/PagedEditor.tsx");
const VUE_PAGED_REF_SRC = path.join(
  repoRoot,
  "packages/vue/src/components/DocxEditor/pagedEditorRef.ts",
);
const CONTRACT_PATH = path.join(repoRoot, "scripts/parity/parity.contract.json");

/**
 * Extract the top-level field/member names from an `export type <Name> = { ... }`
 * block. Tracks brace depth from the opening `{` so nested object literals and
 * sibling type declarations are ignored; skips JSDoc/line comments so braces and
 * `{{template}}` tokens inside prose never perturb the depth counter. Fields are
 * matched at exactly 2-space indent (the repo's field indentation).
 */
function extractTypeMembers(sourceText, typeName) {
  const lines = sourceText.split("\n");
  const startRe = new RegExp(`^export type ${typeName}\\b.*\\{`);
  const startIdx = lines.findIndex((l) => startRe.test(l));
  if (startIdx === -1) return null;

  const members = new Set();
  let depth = 0;
  let inBlockComment = false;

  for (let i = startIdx; i < lines.length; i++) {
    const raw = lines[i];
    const trimmed = raw.trim();

    // Comment handling: skip JSDoc/block comments wholesale so their braces and
    // `{{...}}` tokens do not move the brace depth.
    if (inBlockComment) {
      if (trimmed.includes("*/")) inBlockComment = false;
      continue;
    }
    if (trimmed.startsWith("/*")) {
      if (!trimmed.includes("*/")) inBlockComment = true;
      continue;
    }
    if (trimmed.startsWith("*") || trimmed.startsWith("//")) continue;

    // Strip a trailing line comment before counting braces.
    const code = raw.split("//")[0];

    for (const ch of code) {
      if (ch === "{") depth++;
      else if (ch === "}") depth--;
    }

    // Only pick up top-level fields (depth 1, exactly 2-space indent). A field
    // whose value opens a nested object (e.g. `getContentControls: (...) => {`)
    // is still captured here because the match runs before depth returns.
    const match = /^ {2}(\w+)\??\s*:/.exec(code);
    if (match && depth >= 1) members.add(match[1]);

    if (depth === 0 && i > startIdx) break;
  }

  return members;
}

function readJson(p) {
  try {
    return JSON.parse(fs.readFileSync(p, "utf8"));
  } catch (e) {
    if (e instanceof SyntaxError) {
      console.error(`Malformed JSON in ${p}:`);
      console.error(`  ${e.message}`);
      process.exit(1);
    }
    throw e;
  }
}

const SECTION_SCHEMA = {
  props: { paired: "array", deferredInVue: "object" },
  ref: { paired: "array", deferredInVue: "object" },
  pagedRef: { paired: "array", deferredInVue: "object" },
};

function validateContractShape(contract) {
  const errors = [];
  for (const [top, buckets] of Object.entries(SECTION_SCHEMA)) {
    const section = contract[top];
    if (!section || typeof section !== "object") {
      errors.push(`Missing or invalid top-level key: ${top}`);
      continue;
    }
    for (const [bucket, type] of Object.entries(buckets)) {
      const value = section[bucket];
      const ok =
        type === "array"
          ? Array.isArray(value)
          : typeof value === "object" && value !== null && !Array.isArray(value);
      if (!ok)
        errors.push(
          `contract.${top}.${bucket} must be ${type === "array" ? "an array" : "an object"}`,
        );
    }
    // Every member must live in exactly one bucket within its section.
    const seen = new Set();
    for (const bucket of Object.keys(buckets)) {
      const value = section[bucket];
      const keys = Array.isArray(value) ? value : Object.keys(value ?? {});
      for (const k of keys) {
        if (seen.has(k)) errors.push(`contract.${top}: '${k}' appears in more than one bucket`);
        seen.add(k);
      }
    }
  }
  return errors;
}

/**
 * Apply one section (props or ref) of the contract against the parsed member
 * sets. Pushes human-readable issues onto `issues`.
 */
function checkSection(kind, section, reactMembers, vueMembers, issues) {
  const paired = section.paired;
  const deferred = Object.keys(section.deferredInVue);
  const classified = new Set([...paired, ...deferred]);

  // paired must exist on both adapters.
  for (const k of paired) {
    if (!reactMembers.has(k)) issues.push(`${kind} paired '${k}' missing from React`);
    if (!vueMembers.has(k)) issues.push(`${kind} paired '${k}' missing from Vue`);
  }
  // deferredInVue is a React feature the Vue implementation stubs; the name must
  // exist on React. (It also exists on Vue's type surface — that is expected and
  // not asserted, since the deferral is at the implementation level.)
  for (const k of deferred) {
    if (!reactMembers.has(k))
      issues.push(`${kind} deferredInVue '${k}' missing from React (contract stale)`);
  }
  // Any adapter member not classified is drift.
  for (const k of reactMembers) {
    if (!classified.has(k))
      issues.push(`${kind} '${k}' in React is not declared in the parity contract`);
  }
  for (const k of vueMembers) {
    if (!classified.has(k))
      issues.push(`${kind} '${k}' in Vue is not declared in the parity contract`);
  }
}

function main() {
  for (const f of [
    REACT_PROPS_SRC,
    VUE_TYPES_SRC,
    REACT_PAGED_REF_SRC,
    VUE_PAGED_REF_SRC,
    CONTRACT_PATH,
  ]) {
    if (!fs.existsSync(f)) {
      console.error(`Missing required file: ${f}`);
      process.exit(1);
    }
  }

  const reactSrc = fs.readFileSync(REACT_PROPS_SRC, "utf8");
  const vueSrc = fs.readFileSync(VUE_TYPES_SRC, "utf8");
  const reactPagedRefSrc = fs.readFileSync(REACT_PAGED_REF_SRC, "utf8");
  const vuePagedRefSrc = fs.readFileSync(VUE_PAGED_REF_SRC, "utf8");
  const contract = readJson(CONTRACT_PATH);

  const shapeErrors = validateContractShape(contract);
  if (shapeErrors.length > 0) {
    console.error("Parity contract has structural errors:");
    for (const e of shapeErrors) console.error(`  - ${e}`);
    process.exit(1);
  }

  const sources = [
    ["React DocxEditorProps", reactSrc, "DocxEditorProps"],
    ["Vue DocxEditorProps", vueSrc, "DocxEditorProps"],
    ["React DocxEditorRef", reactSrc, "DocxEditorRef"],
    ["Vue DocxEditorRef", vueSrc, "DocxEditorRef"],
    ["React PagedEditorRef", reactPagedRefSrc, "PagedEditorRef"],
    ["Vue PagedEditorRef", vuePagedRefSrc, "PagedEditorRef"],
  ];
  const parsed = {};
  for (const [label, src, typeName] of sources) {
    const members = extractTypeMembers(src, typeName);
    if (!members || members.size === 0) {
      console.error(`Could not locate ${typeName} members for ${label}`);
      process.exit(1);
    }
    parsed[label] = members;
  }

  const issues = [];
  checkSection(
    "PROP",
    contract.props,
    parsed["React DocxEditorProps"],
    parsed["Vue DocxEditorProps"],
    issues,
  );
  checkSection(
    "REF",
    contract.ref,
    parsed["React DocxEditorRef"],
    parsed["Vue DocxEditorRef"],
    issues,
  );
  checkSection(
    "PAGED REF",
    contract.pagedRef,
    parsed["React PagedEditorRef"],
    parsed["Vue PagedEditorRef"],
    issues,
  );

  console.log(`Parity contract: scripts/parity/parity.contract.json (v${contract.version})`);
  console.log(`  React DocxEditorProps: ${parsed["React DocxEditorProps"].size} fields`);
  console.log(`  Vue   DocxEditorProps: ${parsed["Vue DocxEditorProps"].size} fields`);
  console.log(`  React DocxEditorRef:   ${parsed["React DocxEditorRef"].size} members`);
  console.log(`  Vue   DocxEditorRef:   ${parsed["Vue DocxEditorRef"].size} members`);
  console.log(`  React PagedEditorRef:  ${parsed["React PagedEditorRef"].size} members`);
  console.log(`  Vue   PagedEditorRef:  ${parsed["Vue PagedEditorRef"].size} members`);
  console.log(`  Paired props:          ${contract.props.paired.length}`);
  console.log(`  Deferred-in-Vue props: ${Object.keys(contract.props.deferredInVue).length}`);
  console.log(`  Paired ref members:    ${contract.ref.paired.length}`);
  console.log(`  Deferred-in-Vue refs:  ${Object.keys(contract.ref.deferredInVue).length}`);
  console.log(`  Paired paged refs:     ${contract.pagedRef.paired.length}`);
  console.log(`  Deferred paged refs:  ${Object.keys(contract.pagedRef.deferredInVue).length}`);

  if (issues.length > 0) {
    console.error(`\nParity drift: ${issues.length} issue${issues.length === 1 ? "" : "s"}`);
    for (const issue of issues) console.error(`  - ${issue}`);
    console.error("\nFix: update scripts/parity/parity.contract.json to acknowledge the change,");
    console.error("then commit the contract alongside the adapter change.");
    process.exit(1);
  }

  console.log("\nParity check passed.");
}

main();
