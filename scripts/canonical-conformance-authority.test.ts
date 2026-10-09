import { expect, test } from "bun:test";
import { readFileSync } from "node:fs";
import path from "node:path";
import ts from "typescript";

const parse = (source: string) =>
  ts.createSourceFile("authority.ts", source, ts.ScriptTarget.Latest, true, ts.ScriptKind.TS);

const fixedAuthority = (call: ts.CallExpression, authority: string) => {
  const options = call.arguments.at(0);
  if (!options || !ts.isObjectLiteralExpression(options)) return false;
  const property = options.properties.at(-1);
  return (
    property !== undefined &&
    ts.isPropertyAssignment(property) &&
    ts.isIdentifier(property.name) &&
    property.name.text === "authority" &&
    ts.isStringLiteral(property.initializer) &&
    property.initializer.text === authority
  );
};

const standingFactoryIsCanonical = (source: string) => {
  let found = false;
  const visit = (node: ts.Node) => {
    if (
      ts.isVariableDeclaration(node) &&
      ts.isIdentifier(node.name) &&
      node.name.text === "runConformanceCase"
    ) {
      const factory = node.initializer;
      found = Boolean(
        factory &&
        ts.isArrowFunction(factory) &&
        factory.parameters.length === 1 &&
        ts.isCallExpression(factory.body) &&
        ts.isIdentifier(factory.body.expression) &&
        factory.body.expression.text === "runCaseWithAuthority" &&
        fixedAuthority(factory.body, "canonical"),
      );
    }
    ts.forEachChild(node, visit);
  };
  visit(parse(source));
  return found;
};

const randomBrowserAuthorityProblems = (source: string) => {
  const problems: string[] = [];
  let calls = 0;
  const visit = (node: ts.Node, randomLane = false) => {
    const inLane =
      randomLane || (ts.isForOfStatement(node) && node.expression.getText() === "config.seeds");
    if (inLane && ts.isCallExpression(node)) {
      if (ts.isIdentifier(node.expression) && node.expression.text === "runMode") {
        calls += 1;
        if (!fixedAuthority(node, "canonical")) problems.push("random browser authority");
      }
      if (
        ts.isPropertyAccessExpression(node.expression) &&
        node.expression.getText() === "test.fail"
      )
        problems.push("canonical random lane accepts a legacy failure");
    }
    ts.forEachChild(node, (child) => visit(child, inLane));
  };
  visit(parse(source));
  if (calls === 0) problems.push("missing standing browser calls");
  return problems;
};

test("standing conformance cannot select legacy authority through its public factory", () => {
  const source = readFileSync(
    new URL("../packages/core/src/__tests__/editorCommandConformance.ts", import.meta.url),
    "utf8",
  );
  expect(standingFactoryIsCanonical(source)).toBe(true);
  expect(
    standingFactoryIsCanonical(
      source.replaceAll('authority: "canonical"', 'authority: "prosemirror"'),
    ),
  ).toBe(false);
  expect(standingFactoryIsCanonical("const runConformanceCase = () => legacy();")).toBe(false);
  expect(
    standingFactoryIsCanonical(
      'const runConformanceCase = (options) => runCaseWithAuthority({ authority: "canonical", ...options });',
    ),
  ).toBe(false);
});

test("random browser conformance is strict canonical; legacy acceptance stays in explicit replay", () => {
  const source = readFileSync(
    new URL("../tests/visual/browser-input-fuzz.interactions.spec.ts", import.meta.url),
    "utf8",
  );
  expect(randomBrowserAuthorityProblems(source)).toEqual([]);
  const lane = 'for (const seed of config.seeds) { runMode({ authority: "canonical" }); }';
  expect(randomBrowserAuthorityProblems(lane)).toEqual([]);
  expect(randomBrowserAuthorityProblems(lane.replace('"canonical"', '"prosemirror"'))).toHaveLength(
    1,
  );
  expect(randomBrowserAuthorityProblems(lane.replace("});", "}); test.fail(true);"))).toHaveLength(
    1,
  );
  expect(randomBrowserAuthorityProblems("for (const seed of config.seeds) {}")).toHaveLength(1);
});

// Representation readers and execution owners are allowed by exact export name.
// Refusal, eligibility and planner decisions cannot become expectation inputs
// through another export from a mixed driver module.
const CONFORMANCE_PRODUCTION_IMPORTS = new Map<string, readonly string[]>([
  // The planner may supply tie ORDER only; expected results stay independent.
  ["packages/core/src/ai-edits/batch-claims", ["compareBatchTieOrder"]],
  ["packages/core/src/docx/paragraphPropertySource", ["cloneDocumentWithParagraphPropertySources"]],
  [
    "packages/core/src/controller/hiddenEditorManager",
    ["createHiddenEditorManager", "CanonicalSessionRefusalError"],
  ],
  [
    "packages/core/src/controller/canonicalSession",
    ["createCanonicalSession", "publishCanonicalProjection"],
  ],
  ["packages/core/src/prosemirror/executeEditorCommand", ["executeEditorCommand"]],
  ["packages/core/src/prosemirror/rangeAnchorAttrs", ["expectRangeAnchorAttrs"]],
  ["packages/core/src/types/canonicalCapabilities", ["CANONICAL_GAP", "CANONICAL_CAPABILITIES"]],
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
    ["acceptAllChanges", "acceptChange", "addCommentMark", "rejectAllChanges", "rejectChange"],
  ],
  ["packages/core/src/prosemirror/commands/formatting", ["clearFormatting"]],
  ["packages/core/src/prosemirror/commands/pageBreak", ["insertPageBreak"]],
  ["packages/core/src/prosemirror/conversion/fromProseDoc", ["fromProseDoc"]],
  ["packages/core/src/prosemirror/plugins/suggestionMode", ["deleteSelectionAsSuggestion"]],
  ["packages/core/src/prosemirror/styles/styleResolver", ["createStyleResolver"]],
  ["packages/core/src/docx/noteReferenceMark", ["createNote"]],
  ["packages/core/src/docx/parser", ["parseDocx"]],
  ["packages/docx-core/src/validate/docx", ["validateDocxPackage"]],
]);

const expectationImportProblems = (source: string, owner: string) =>
  parse(source).statements.flatMap((node) => {
    if (!ts.isImportDeclaration(node) || !ts.isStringLiteral(node.moduleSpecifier)) return [];
    const clause = node.importClause;
    if (clause?.isTypeOnly) return [];
    const module = node.moduleSpecifier.text;
    if (!module.startsWith(".") && !module.startsWith("@stll/")) return [];
    const resolved = path.posix.normalize(path.posix.join(path.posix.dirname(owner), module));
    if (!resolved.startsWith("packages/") && !module.startsWith("@stll/")) return [];
    if (!clause) return [`${module}: side-effect import`];
    const allowed = CONFORMANCE_PRODUCTION_IMPORTS.get(resolved) ?? [];
    const problems: string[] = [];
    if (clause.name) problems.push(`${module}: default import`);
    const bindings = clause.namedBindings;
    if (bindings && ts.isNamespaceImport(bindings)) problems.push(`${module}: namespace import`);
    if (bindings && ts.isNamedImports(bindings))
      for (const binding of bindings.elements) {
        if (binding.isTypeOnly) continue;
        const name = (binding.propertyName ?? binding.name).text;
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
    const source = readFileSync(new URL(`../${owner}`, import.meta.url), "utf8");
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
  expect(expectationImportProblems('import { Slice } from "prosemirror-model";', owner)).toEqual(
    [],
  );
  expect(
    expectationImportProblems(
      'import "../packages/core/src/controller/canonicalClipboard";',
      owner,
    ),
  ).toHaveLength(1);
});
