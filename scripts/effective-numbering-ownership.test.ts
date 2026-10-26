import { expect, test } from "bun:test";
import { execFileSync } from "node:child_process";
import { readFileSync } from "node:fs";
import path from "node:path";

import {
  bindingElements,
  collectSyntaxTree,
  isBlock,
  isFunctionLike,
  type Node,
  nodeText,
  parameterParts,
  parametersOf,
  parseLikeTypeScript,
  walkAst,
} from "./lib/oxc-ast";

const ROOT = path.resolve(import.meta.dir, "..");
const OWNER = "packages/core/src/prosemirror/numberingAttr.ts";
const MEMBERSHIP_READERS = new Set([
  "resolveListState",
  "resolveParagraphNumbering",
  "mergeParagraphNumbering",
  "paragraphNumberingReferenceId",
  "paragraphNumberingLevel",
]);
const EFFECTIVE_READERS = new Set([
  "effectiveParagraphNumbering",
  "effectiveParagraphNumberingReference",
]);
const LEVEL_OWNER_ALLOWLIST = new Set([
  "packages/core/src/prosemirror/paragraphList.ts",
  "packages/core/src/docx/numberingParser.ts",
  "packages/core/src/prosemirror/conversion/fromProseDoc.ts",
  "packages/core/src/prosemirror/listInstanceReferences.ts",
  "packages/core/src/prosemirror/listMarker.ts",
  "packages/core/src/prosemirror/commands/propertyChangeScope.ts",
  "packages/core/src/ai-edits/apply.ts",
  "packages/core/src/ai-edits/headless.ts",
  "packages/core/src/controller/canonicalStructure.ts",
  "packages/core/src/docx/listMarkerComputation.ts",
  "packages/core/src/docx/listModelNormalization.ts",
  "packages/core/src/docx/listRendering.ts",
  "packages/core/src/docx/paragraphParser.ts",
  "packages/core/src/prosemirror/listNumbering.ts",
  "packages/core/src/prosemirror/listState.ts",
  "packages/core/src/prosemirror/storyListNumbering.ts",
  "packages/core/src/prosemirror/styles/resolvedStyleAttrs.ts",
]);

const parseFixture = (source: string, file: string) =>
  parseLikeTypeScript(file, source, "ts").program;

/** `a.b` (TypeScript's property access): the accessed name. */
const accessedName = (node: Node): string | undefined => {
  if (node.type !== "MemberExpression" || node.computed) return undefined;
  return node.property.type === "PrivateIdentifier"
    ? `#${node.property.name}`
    : node.property.name;
};

/** `a["b"]` (TypeScript's element access with a string literal): the key. */
const stringKey = (node: Node): string | undefined =>
  node.type === "MemberExpression" &&
  node.computed &&
  node.property.type === "Literal" &&
  typeof node.property.value === "string"
    ? node.property.value
    : undefined;

const levelOwnerViolations = (
  source: string,
  file = "fixture.ts",
): string[] => {
  const program = parseFixture(source, file);
  const found: string[] = [];
  walkAst(program, (node) => {
    if (
      node.type === "CallExpression" &&
      ((node.callee.type === "Identifier" &&
        node.callee.name === "hasListRendering") ||
        accessedName(node.callee) === "getLevel")
    ) {
      found.push(nodeText(source, node));
    }
    if (
      node.type === "MemberExpression" &&
      accessedName(node) === "levels" &&
      node.object.type === "MemberExpression" &&
      node.object.computed &&
      nodeText(source, node.object.property).includes("abstractNum")
    ) {
      found.push(nodeText(source, node));
    }
  });
  return found;
};

const cachedRenderingDecisionViolations = (
  source: string,
  file = "fixture.ts",
): string[] => {
  if (
    !source.includes("resolveParagraphList") ||
    !source.includes("listRendering")
  )
    return [];
  const program = parseFixture(source, file);
  const { nodes, parentOf } = collectSyntaxTree(program);
  const found: string[] = [];
  const decisionScope = (node: Node): Node => {
    let scope = parentOf(node);
    while (scope && !isFunctionLike(scope) && scope.type !== "Program")
      scope = parentOf(scope);
    return scope ?? program;
  };
  // TypeScript's text of a function: a method's starts at its name, and the
  // file's runs to the end of the text.
  const scopeText = (scope: Node): string => {
    if (scope.type === "Program") return source.slice(scope.start);
    const owner = parentOf(scope);
    return scope.type === "FunctionExpression" &&
      owner !== undefined &&
      (owner.type === "MethodDefinition" ||
        owner.type === "TSAbstractMethodDefinition" ||
        (owner.type === "Property" && owner.value === scope))
      ? nodeText(source, owner)
      : nodeText(source, scope);
  };
  for (const node of nodes) {
    if (
      (accessedName(node) === "listRendering" ||
        stringKey(node) === "listRendering") &&
      scopeText(decisionScope(node)).includes("resolveParagraphList(")
    ) {
      found.push(nodeText(source, node));
    }
  }
  return found;
};

/** `object.property` built for a destructured name; it has no source of its own. */
type SyntheticAccess = {
  readonly type: "SyntheticAccess";
  readonly object: Node;
  readonly property: string;
};
type Expression = Node | SyntheticAccess;

/** Follow local aliases so membership cannot accidentally read only the authored slot. */
const violations = (source: string, file = "fixture.ts"): string[] => {
  // Files without either a source slot or a membership reader cannot cross
  // this boundary. Avoid constructing their syntax trees during discovery.
  if (
    !source.includes("numPr") ||
    ![...MEMBERSHIP_READERS].some((name) => source.includes(name))
  )
    return [];
  const program = parseFixture(source, file);
  const { nodes, parentOf } = collectSyntaxTree(program);
  const scopeOf = (node: Node): Node => {
    let scope = parentOf(node);
    while (
      scope &&
      !isBlock(scope) &&
      scope.type !== "Program" &&
      !isFunctionLike(scope)
    )
      scope = parentOf(scope);
    return scope ?? program;
  };
  const bindings = new Map<Node, Map<string, Expression | undefined>>();
  const localsOf = (scope: Node): Map<string, Expression | undefined> => {
    const locals =
      bindings.get(scope) ?? new Map<string, Expression | undefined>();
    bindings.set(scope, locals);
    return locals;
  };
  for (const node of nodes) {
    if (node.type === "VariableDeclarator" && node.id.type === "Identifier")
      localsOf(scopeOf(node)).set(node.id.name, node.init ?? undefined);
    // TypeScript declares a catch clause's variable like any other.
    if (node.type === "CatchClause" && node.param?.type === "Identifier")
      localsOf(scopeOf(node.param)).set(node.param.name, undefined);
    if (
      node.type === "VariableDeclarator" &&
      node.id.type === "ObjectPattern" &&
      node.init
    ) {
      const locals = localsOf(scopeOf(node));
      for (const element of bindingElements(node.id)) {
        if (
          element.name.type !== "Identifier" ||
          element.node.type === "RestElement"
        )
          continue;
        const property = element.propertyName;
        if (
          property !== undefined &&
          (property.type !== "Identifier" ||
            (element.node.type === "Property" && element.node.computed))
        )
          continue;
        locals.set(element.name.name, {
          type: "SyntheticAccess",
          object: node.init,
          property: property?.name ?? element.name.name,
        });
      }
    }
    if (isFunctionLike(node)) {
      const locals = localsOf(node);
      for (const parameter of parametersOf(node)) {
        const { name } = parameterParts(parameter);
        if (name.type === "Identifier") locals.set(name.name, undefined);
      }
    }
  }
  const initializerOf = (node: Node & { name: string }) => {
    let scope: Node | undefined = node;
    while (scope) {
      const locals = bindings.get(scope);
      if (locals?.has(node.name)) return locals.get(node.name);
      scope = parentOf(scope);
    }
    return undefined;
  };
  const importNames = new Map<string, string>();
  for (const node of nodes)
    if (node.type === "ImportSpecifier")
      importNames.set(
        node.local.name,
        node.imported.type === "Literal"
          ? node.imported.value
          : node.imported.name,
      );
  const propertyAccessName = (node: Expression): string | undefined =>
    node.type === "SyntheticAccess" ? node.property : accessedName(node);
  const callableName = (
    node: Expression,
    seen = new Set<string>(),
  ): string | undefined => {
    const accessed = propertyAccessName(node);
    if (accessed !== undefined) return accessed;
    if (node.type !== "Identifier") return undefined;
    if (seen.has(node.name)) return undefined;
    const initializer = initializerOf(node);
    if (initializer) {
      const next = new Set(seen);
      next.add(node.name);
      return callableName(initializer, next);
    }
    return importNames.get(node.name) ?? node.name;
  };
  const isAttrs = (node: Expression, seen = new Set<string>()): boolean => {
    if (node.type === "Identifier") {
      if (/attrs$/iu.test(node.name)) return true;
      const initializer = initializerOf(node);
      if (initializer && !seen.has(node.name)) {
        const next = new Set(seen);
        next.add(node.name);
        return isAttrs(initializer, next);
      }
    }
    return (
      propertyAccessName(node) === "attrs" ||
      (node.type === "CallExpression" &&
        ["expectParagraphAttrs", "mergeParagraphAttrs"].includes(
          callableName(node.callee) ?? "",
        ))
    );
  };
  const readsAuthoredSlot = (
    node: Expression,
    seen = new Set<string>(),
  ): boolean => {
    if (
      node.type === "CallExpression" &&
      EFFECTIVE_READERS.has(callableName(node.callee) ?? "")
    )
      return false;
    if (
      (node.type === "SyntheticAccess" || node.type === "MemberExpression") &&
      (propertyAccessName(node) === "numPr" ||
        (node.type === "MemberExpression" && stringKey(node) === "numPr")) &&
      isAttrs(node.object)
    )
      return true;
    if (node.type === "Identifier" && !seen.has(node.name)) {
      const initializer = initializerOf(node);
      if (initializer) {
        const next = new Set(seen);
        next.add(node.name);
        return readsAuthoredSlot(initializer, next);
      }
    }
    // A synthetic access's own name resolves to nothing; its object is its
    // only child that can.
    if (node.type === "SyntheticAccess")
      return readsAuthoredSlot(node.object, seen);
    let found = false;
    walkAst(node, (child) => {
      if (child === node) return true;
      if (!found) found = readsAuthoredSlot(child, seen);
      return false;
    });
    return found;
  };
  const found: string[] = [];
  for (const node of nodes) {
    if (
      node.type === "CallExpression" &&
      MEMBERSHIP_READERS.has(callableName(node.callee) ?? "") &&
      node.arguments.some((argument) => readsAuthoredSlot(argument))
    )
      found.push(nodeText(source, node));
  }
  return found;
};

test("membership readers require the effective numbering owner through local aliases", () => {
  expect(violations("resolveListState(map, attrs.numPr)")).toHaveLength(1);
  expect(
    violations(
      "const direct = attrs['numPr']; resolveParagraphNumbering(direct)",
    ),
  ).toHaveLength(1);
  expect(
    violations("resolveListState(map, effectiveParagraphNumbering(attrs))"),
  ).toEqual([]);
  expect(
    violations(
      "const effective = effectiveParagraphNumberingReference(attrs); resolveListState(map, effective)",
    ),
  ).toEqual([]);
  expect(
    violations("mergeParagraphNumbering(attrs.numPrFromStyle, attrs.numPr)"),
  ).toHaveLength(1);
  expect(
    violations(
      "const direct = attrs.numPr; function safe(direct) { resolveListState(map, direct); }",
    ),
  ).toEqual([]);
  expect(
    violations(
      "import { resolveListState as resolve } from './listState'; resolve(map, expectParagraphAttrs(node).numPr)",
    ),
  ).toHaveLength(1);
  expect(
    violations("const resolve = resolveListState; resolve(map, attrs.numPr)"),
  ).toHaveLength(1);
  expect(
    violations(
      "const { numPr: direct } = attrs; resolveListState(map, direct)",
    ),
  ).toHaveLength(1);
});

test("production membership readers never consume the authored slot alone", () => {
  const files = execFileSync("git", ["ls-files", "packages/core/src"], {
    cwd: ROOT,
  })
    .toString()
    .trim()
    .split("\n")
    .filter(
      (file) =>
        file.endsWith(".ts") && !file.endsWith(".test.ts") && file !== OWNER,
    );
  const found = files.flatMap((file) =>
    violations(readFileSync(path.join(ROOT, file), "utf8"), file).map(
      (entry) => `${file}: ${entry}`,
    ),
  );
  expect(found).toEqual([]);
});

test("numbering level decisions stay in explicitly owned bridges", () => {
  expect(
    levelOwnerViolations("const shown = hasListRendering(attrs)"),
  ).toHaveLength(1);
  expect(
    levelOwnerViolations("const level = numbering.getLevel(2, 1)"),
  ).toHaveLength(1);
  expect(
    levelOwnerViolations("resolveParagraphList(sources, numbering)"),
  ).toEqual([]);

  const files = execFileSync("git", ["ls-files", "packages/core/src"], {
    cwd: ROOT,
  })
    .toString()
    .trim()
    .split("\n")
    .filter(
      (file) =>
        file.endsWith(".ts") &&
        !file.endsWith(".test.ts") &&
        !file.includes("/layout-bridge/") &&
        !LEVEL_OWNER_ALLOWLIST.has(file),
    );
  const found = files.flatMap((file) =>
    levelOwnerViolations(readFileSync(path.join(ROOT, file), "utf8"), file).map(
      (entry) => `${file}: ${entry}`,
    ),
  );
  expect(found).toEqual([]);
});

test("paragraph list readers do not decide rendering from the cached projection", () => {
  expect(
    cachedRenderingDecisionViolations(
      "const list = resolveParagraphList(sources, numbering); return paragraph.listRendering;",
    ),
  ).toHaveLength(1);
  expect(
    cachedRenderingDecisionViolations(
      "const list = resolveParagraphList(sources, numbering); return list.rendering;",
    ),
  ).toEqual([]);

  const files = execFileSync("git", ["ls-files", "packages/core/src"], {
    cwd: ROOT,
  })
    .toString()
    .trim()
    .split("\n")
    .filter((file) => file.endsWith(".ts") && !file.endsWith(".test.ts"));
  const found = files.flatMap((file) =>
    cachedRenderingDecisionViolations(
      readFileSync(path.join(ROOT, file), "utf8"),
      file,
    ).map((entry) => `${file}: ${entry}`),
  );
  expect(found).toEqual([]);
});
