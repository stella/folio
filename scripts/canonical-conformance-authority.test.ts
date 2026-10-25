import { expect, test } from "bun:test";
import { readFileSync } from "node:fs";
import path from "node:path";

import { type Node, nodeText, parseSourceModule, walkAst } from "./lib/oxc-ast";

const parse = (source: string) =>
  parseSourceModule("authority.ts", source).program;

const fixedAuthority = (call: Node, authority: string) => {
  if (call.type !== "CallExpression") return false;
  const options = call.arguments.at(0);
  if (options?.type !== "ObjectExpression") return false;
  const property = options.properties.at(-1);
  return (
    property?.type === "Property" &&
    property.kind === "init" &&
    !property.method &&
    !property.shorthand &&
    !property.computed &&
    property.key.type === "Identifier" &&
    property.key.name === "authority" &&
    property.value.type === "Literal" &&
    property.value.value === authority
  );
};

const standingFactoryIsCanonical = (source: string) => {
  let found = false;
  walkAst(parse(source), (node) => {
    if (
      node.type === "VariableDeclarator" &&
      node.id.type === "Identifier" &&
      node.id.name === "runConformanceCase"
    ) {
      const factory = node.init;
      found = Boolean(
        factory &&
        factory.type === "ArrowFunctionExpression" &&
        factory.params.length === 1 &&
        factory.body.type === "CallExpression" &&
        factory.body.callee.type === "Identifier" &&
        factory.body.callee.name === "runCaseWithAuthority" &&
        fixedAuthority(factory.body, "canonical"),
      );
    }
  });
  return found;
};

const randomBrowserAuthorityProblems = (source: string) => {
  const problems: string[] = [];
  let calls = 0;
  walkAst(parse(source), (node, ancestors) => {
    const inLane = [...ancestors, node].some(
      (lane) =>
        lane.type === "ForOfStatement" &&
        nodeText(source, lane.right) === "config.seeds",
    );
    if (inLane && node.type === "CallExpression") {
      if (node.callee.type === "Identifier" && node.callee.name === "runMode") {
        calls += 1;
        if (!fixedAuthority(node, "canonical"))
          problems.push("random browser authority");
      }
      if (
        node.callee.type === "MemberExpression" &&
        !node.callee.computed &&
        nodeText(source, node.callee) === "test.fail"
      )
        problems.push("canonical random lane accepts a legacy failure");
    }
  });
  if (calls === 0) problems.push("missing standing browser calls");
  return problems;
};

test("standing conformance cannot select legacy authority through its public factory", () => {
  const source = readFileSync(
    new URL(
      "../packages/core/src/__tests__/editorCommandConformance.ts",
      import.meta.url,
    ),
    "utf8",
  );
  expect(standingFactoryIsCanonical(source)).toBe(true);
  expect(
    standingFactoryIsCanonical(
      source.replaceAll('authority: "canonical"', 'authority: "prosemirror"'),
    ),
  ).toBe(false);
  expect(
    standingFactoryIsCanonical("const runConformanceCase = () => legacy();"),
  ).toBe(false);
  expect(
    standingFactoryIsCanonical(
      'const runConformanceCase = (options) => runCaseWithAuthority({ authority: "canonical", ...options });',
    ),
  ).toBe(false);
});

test("random browser conformance is strict canonical; legacy acceptance stays in explicit replay", () => {
  const source = readFileSync(
    new URL(
      "../tests/visual/browser-input-fuzz.interactions.spec.ts",
      import.meta.url,
    ),
    "utf8",
  );
  expect(randomBrowserAuthorityProblems(source)).toEqual([]);
  const lane =
    'for (const seed of config.seeds) { runMode({ authority: "canonical" }); }';
  expect(randomBrowserAuthorityProblems(lane)).toEqual([]);
  expect(
    randomBrowserAuthorityProblems(
      lane.replace('"canonical"', '"prosemirror"'),
    ),
  ).toHaveLength(1);
  expect(
    randomBrowserAuthorityProblems(lane.replace("});", "}); test.fail(true);")),
  ).toHaveLength(1);
  expect(
    randomBrowserAuthorityProblems("for (const seed of config.seeds) {}"),
  ).toHaveLength(1);
});

// Representation readers and execution owners are allowed by exact export name.
// Refusal, eligibility and planner decisions cannot become expectation inputs
// through another export from a mixed driver module.
const CONFORMANCE_PRODUCTION_IMPORTS = new Map<string, readonly string[]>([
  // The planner may supply tie ORDER only; expected results stay independent.
  ["packages/core/src/ai-edits/batch-claims", ["compareBatchTieOrder"]],
  [
    "packages/core/src/docx/paragraphPropertySource",
    ["cloneDocumentWithParagraphPropertySources"],
  ],
  [
    "packages/core/src/controller/hiddenEditorManager",
    ["createHiddenEditorManager", "CanonicalSessionRefusalError"],
  ],
  [
    "packages/core/src/controller/canonicalSession",
    ["createCanonicalSession", "publishCanonicalProjection"],
  ],
  [
    "packages/core/src/prosemirror/executeEditorCommand",
    ["executeEditorCommand"],
  ],
  [
    "packages/core/src/prosemirror/rangeAnchorAttrs",
    ["expectRangeAnchorAttrs"],
  ],
  [
    "packages/core/src/types/canonicalCapabilities",
    ["CANONICAL_GAP", "CANONICAL_CAPABILITIES"],
  ],
  ["packages/core/src/docx/canonicalSave", ["serializeCanonicalSave"]],
  [
    "packages/core/src/__tests__/editorHarness",
    [
      "keyboardEventFor",
      "HARNESS_AUTHOR",
      "createHarnessState",
      "EDITOR_MODES",
      "harnessRuntimeManager",
      "HeadlessEditorView",
      "modelMarkdown",
      "parseShapeDocument",
      "placeSelection",
      "readBack",
      "resolveAllChanges",
      "saveHarnessState",
      "summarizeEffectiveParagraphs",
      "summarizeState",
      "textblocks",
    ],
  ],
  ["packages/core/src/prosemirror/schema", ["singletonManager"]],
  ["packages/core/src/docx/modelValidation", ["assertValidFolioDocumentModel"]],
  ["packages/core/src/prosemirror/conversion/toProseDoc", ["toProseDoc"]],
  ["packages/core/src/docx/numberingParser", ["getCachedNumberingMap"]],
  [
    "packages/core/src/prosemirror/commands/comments",
    [
      "acceptAllChanges",
      "acceptChange",
      "addCommentMark",
      "rejectAllChanges",
      "rejectChange",
    ],
  ],
  ["packages/core/src/prosemirror/commands/formatting", ["clearFormatting"]],
  ["packages/core/src/prosemirror/commands/pageBreak", ["insertPageBreak"]],
  ["packages/core/src/prosemirror/conversion/fromProseDoc", ["fromProseDoc"]],
  [
    "packages/core/src/prosemirror/plugins/suggestionMode",
    ["deleteSelectionAsSuggestion"],
  ],
  [
    "packages/core/src/prosemirror/styles/styleResolver",
    ["createStyleResolver"],
  ],
  ["packages/core/src/docx/noteReferenceMark", ["createNote"]],
  ["packages/core/src/docx/parser", ["parseDocx"]],
  ["packages/docx-core/src/validate/docx", ["validateDocxPackage"]],
]);

const expectationImportProblems = (source: string, owner: string) =>
  parse(source).body.flatMap((node) => {
    if (node.type !== "ImportDeclaration") return [];
    if (node.importKind === "type") return [];
    const module = node.source.value;
    if (!module.startsWith(".") && !module.startsWith("@stll/")) return [];
    const resolved = path.posix.normalize(
      path.posix.join(path.posix.dirname(owner), module),
    );
    if (!resolved.startsWith("packages/") && !module.startsWith("@stll/"))
      return [];
    // `import {} from` still has an import clause; only `import "x"` has none.
    const hasClause =
      node.specifiers.length > 0 ||
      source.slice(node.start, node.source.start).includes("{");
    if (!hasClause) return [`${module}: side-effect import`];
    const allowed = CONFORMANCE_PRODUCTION_IMPORTS.get(resolved) ?? [];
    const problems: string[] = [];
    for (const binding of node.specifiers) {
      if (binding.type === "ImportDefaultSpecifier")
        problems.push(`${module}: default import`);
      if (binding.type === "ImportNamespaceSpecifier")
        problems.push(`${module}: namespace import`);
      if (binding.type !== "ImportSpecifier" || binding.importKind === "type")
        continue;
      const name =
        binding.imported.type === "Identifier"
          ? binding.imported.name
          : binding.imported.value;
      if (!allowed.includes(name)) problems.push(`${module}: ${name}`);
    }
    return problems;
  });

test("conformance expectations cannot import production decision code", () => {
  for (const owner of [
    "test/canonicalEditorHarness.ts",
    "test/canonical-conformance-refusals.ts",
    "test/canonical-refusal-rows.ts",
    "packages/core/src/__tests__/editorCommandConformance.ts",
    "tests/visual/canonicalBrowserHistoryOracle.ts",
  ]) {
    const source = readFileSync(
      new URL(`../${owner}`, import.meta.url),
      "utf8",
    );
    expect(expectationImportProblems(source, owner)).toEqual([]);
  }
});

test("the expectation import guard catches predicates in new and mixed modules", () => {
  const owner = "test/canonicalEditorHarness.ts";
  expect(
    expectationImportProblems(
      'import { compareBatchTieOrder } from "../packages/core/src/ai-edits/batch-claims";',
      owner,
    ),
  ).toEqual([]);
  expect(
    expectationImportProblems(
      'import { planBatch } from "../packages/core/src/ai-edits/batch-claims";',
      owner,
    ),
  ).toHaveLength(1);
  for (const module of [
    "canonicalClipboard",
    "canonicalCommands",
    "canonicalStructure",
    "eligibility",
  ])
    expect(
      expectationImportProblems(
        `import { predicate } from "../packages/core/src/controller/${module}";`,
        owner,
      ),
    ).toHaveLength(1);
  expect(
    expectationImportProblems(
      'import { createCanonicalSession, isEligible as allowed } from "../packages/core/src/controller/canonicalSession";',
      owner,
    ),
  ).toHaveLength(1);
  expect(
    expectationImportProblems(
      'import * as decisions from "../packages/core/src/controller/canonicalSession";',
      owner,
    ),
  ).toHaveLength(1);
  expect(
    expectationImportProblems(
      'import predicate from "../packages/core/src/controller/canonicalClipboard";',
      owner,
    ),
  ).toHaveLength(1);
  expect(
    expectationImportProblems(
      'import { createCanonicalSession } from "../packages/core/src/controller/canonicalSession";',
      owner,
    ),
  ).toEqual([]);
  expect(
    expectationImportProblems(
      'import type { CanonicalSession } from "../packages/core/src/controller/canonicalSession";',
      owner,
    ),
  ).toEqual([]);
  expect(
    expectationImportProblems(
      'import { Slice } from "prosemirror-model";',
      owner,
    ),
  ).toEqual([]);
  expect(
    expectationImportProblems(
      'import "../packages/core/src/controller/canonicalClipboard";',
      owner,
    ),
  ).toHaveLength(1);
});
