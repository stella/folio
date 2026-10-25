import { expect, setDefaultTimeout, test } from "bun:test";
import path from "node:path";

import type { CanonicalCommandIntent } from "../packages/core/src/prosemirror/canonicalCommands";
import { scanTestTimeout } from "../test/scan-timeout";
import {
  type Node,
  nodeText,
  parseSourceModule,
  type Program,
  walkAst,
} from "./lib/oxc-ast";

setDefaultTimeout(scanTestTimeout(30_000));

const root = path.resolve(import.meta.dir, "..");
const producers = {
  formatRun: [
    "canonicalCommands.ts#canonicalRunFormatting",
    "extensions/marks/markUtils.ts#canonicalSetMarkFormatting",
    "extensions/marks/markUtils.ts#clearFormatting",
  ],
  setHyperlink: ["extensions/marks/HyperlinkExtension.ts#setHyperlink"],
  removeHyperlink: ["extensions/marks/HyperlinkExtension.ts#removeHyperlink"],
  insertHyperlink: ["extensions/marks/HyperlinkExtension.ts#insertHyperlink"],
  insertBreak: ["commands/pageBreak.ts#insertPageBreak"],
} as const satisfies Record<
  Extract<CanonicalCommandIntent, { from: number; to: number }>["type"],
  readonly string[]
>;

/** A property written as `name: value`, as opposed to a shorthand, method or accessor. */
const isPropertyAssignment = (node: Node) =>
  node.type === "Property" &&
  node.kind === "init" &&
  !node.method &&
  !node.shorthand;

/** The written name of a property key; a computed key keeps its brackets. */
const keyText = (text: string, node: Node): string | undefined => {
  if (node.type !== "Property") return undefined;
  const key = nodeText(text, node.key);
  return node.computed ? `[${key}]` : key;
};

const typeOf = (text: string, node: Node) => {
  if (node.type !== "ObjectExpression") return undefined;
  for (const property of node.properties) {
    if (
      !isPropertyAssignment(property) ||
      property.type !== "Property" ||
      keyText(text, property) !== "type"
    )
      continue;
    const value =
      property.value.type === "TSAsExpression"
        ? property.value.expression
        : property.value;
    if (value.type === "Literal" && typeof value.value === "string")
      return value.value;
  }
  return undefined;
};

/** The nearest named variable or function declaration on `lineage` (root first, node last). */
const producerDeclaration = (lineage: readonly Node[]): Node | undefined =>
  lineage.findLast(
    (node) =>
      (node.type === "VariableDeclarator" && node.id.type === "Identifier") ||
      node.type === "FunctionDeclaration" ||
      node.type === "TSDeclareFunction",
  );

/** A binding as written, without the type annotation oxc includes in its span. */
const bindingText = (text: string, binding: Node) => {
  const annotation =
    binding.type === "Identifier" ||
    binding.type === "ObjectPattern" ||
    binding.type === "ArrayPattern"
      ? binding.typeAnnotation
      : undefined;
  return annotation
    ? text.slice(binding.start, annotation.start).trimEnd()
    : nodeText(text, binding);
};

const declarationName = (text: string, declaration: Node | undefined) => {
  if (
    declaration?.type === "VariableDeclarator" ||
    declaration?.type === "FunctionDeclaration" ||
    declaration?.type === "TSDeclareFunction"
  )
    return declaration.id ? bindingText(text, declaration.id) : undefined;
  return undefined;
};

const rangeInputRoot = (text: string, lineage: readonly Node[]): Node => {
  const node = lineage.at(-1);
  if (node === undefined) throw new TypeError("Range input has no node");
  for (let index = lineage.length - 1; index > 0; index -= 1) {
    const current = lineage[index];
    const parent = lineage[index - 1];
    if (
      current !== undefined &&
      parent?.type === "CallExpression" &&
      nodeText(text, parent.callee) === "withCanonicalCommand" &&
      parent.arguments.at(1) === current
    )
      return current;
  }
  const declaration = producerDeclaration(lineage);
  if (declaration?.type === "VariableDeclarator")
    return declaration.init ?? declaration;
  return declaration ?? node;
};

const memberName = (input: Node) => {
  const node = input.type === "ChainExpression" ? input.expression : input;
  if (node.type !== "MemberExpression") return undefined;
  if (!node.computed)
    return node.property.type === "PrivateIdentifier"
      ? `#${node.property.name}`
      : node.property.name;
  if (
    node.property.type === "Literal" &&
    typeof node.property.value === "string"
  )
    return node.property.value;
  return undefined;
};

/** The written binding key of a destructured property (`from` in `from: left`). */
const bindingKey = (text: string, binding: Node) =>
  binding.type === "RestElement"
    ? nodeText(text, binding.argument)
    : keyText(text, binding);

/** The local binding of a destructured property, before any default value. */
const bindingTarget = (binding: Node) => {
  if (binding.type === "RestElement") return binding.argument;
  if (binding.type !== "Property") return undefined;
  return binding.value.type === "AssignmentPattern"
    ? binding.value.left
    : binding.value;
};

/** Follow descriptor helpers, regardless of their names or selection binding syntax. */
const rawRangeInputs = (
  text: string,
  program: Program,
  roots: readonly Node[],
) => {
  const helpers = new Map<string, Node[]>();
  const declarations: Node[] = [];
  walkAst(program, (node) => {
    if (node.type === "VariableDeclarator") {
      declarations.push(node);
      if (
        node.id.type === "Identifier" &&
        node.init &&
        (node.init.type === "ArrowFunctionExpression" ||
          node.init.type === "FunctionExpression")
      ) {
        const entries = helpers.get(node.id.name) ?? [];
        entries.push(node.init);
        helpers.set(node.id.name, entries);
      }
    }
    if (
      (node.type === "FunctionDeclaration" ||
        node.type === "TSDeclareFunction") &&
      node.id
    ) {
      const entries = helpers.get(node.id.name) ?? [];
      entries.push(node);
      helpers.set(node.id.name, entries);
    }
  });
  const aliases = new Set<string>();
  const isSelection = (node: Node) => {
    const value = node.type === "ChainExpression" ? node.expression : node;
    return (
      memberName(value) === "selection" ||
      (value.type === "Identifier" && aliases.has(value.name))
    );
  };
  let previousSize = -1;
  while (aliases.size !== previousSize) {
    previousSize = aliases.size;
    for (const declaration of declarations) {
      if (declaration.type !== "VariableDeclarator" || !declaration.init)
        continue;
      if (declaration.id.type === "Identifier" && isSelection(declaration.init))
        aliases.add(declaration.id.name);
      if (declaration.id.type === "ObjectPattern") {
        for (const binding of declaration.id.properties) {
          const target = bindingTarget(binding);
          if (
            bindingKey(text, binding) === "selection" &&
            target?.type === "Identifier"
          )
            aliases.add(target.name);
        }
      }
    }
  }
  const failures: string[] = [];
  const visited = new Set<Node>();
  const visit = (start: Node): void => {
    walkAst(start, (node) => {
      if (visited.has(node)) return false;
      visited.add(node);
      if (
        node.type === "MemberExpression" &&
        ["from", "to"].includes(memberName(node) ?? "") &&
        isSelection(node.object)
      )
        failures.push(nodeText(text, node));
      if (
        node.type === "VariableDeclarator" &&
        node.init &&
        isSelection(node.init) &&
        node.id.type === "ObjectPattern"
      ) {
        for (const binding of node.id.properties) {
          if (["from", "to"].includes(bindingKey(text, binding) ?? ""))
            failures.push(nodeText(text, binding));
        }
      }
      if (node.type === "CallExpression" && node.callee.type === "Identifier") {
        for (const helper of helpers.get(node.callee.name) ?? []) visit(helper);
      }
      return true;
    });
  };
  for (const rootNode of roots) visit(rootNode);
  return failures;
};

test("range guard follows the called helper across raw selection binding shapes", () => {
  const bindings = [
    "return { from: state.selection.from, to: state.selection.to };",
    "const { from, to, empty } = state.selection; return { from, to };",
    "const { from: left, to: right } = state.selection; return { from: left, to: right };",
    "const { selection: selected } = state; const { from, to } = selected; return { from, to };",
    "const selected = state['selection']; return { from: selected['from'], to: selected['to'] };",
    "return canonicalSelectionRange(state);",
  ];
  for (const [index, binding] of bindings.entries()) {
    const text = `
      import { canonicalSelectionRange } from './canonicalSelectionRange';
      const canonicalRemovalRange = (state) => { ${binding} };
      const descriptor = (state) => {
        const range = canonicalRemovalRange(state);
        return [{ type: 'removeHyperlink', ...range }];
      };
      withCanonicalCommand(raw, descriptor);
    `;
    const { program } = parseSourceModule("fixture.ts", text);
    const descriptor = program.body.find(
      (statement) =>
        statement.type === "VariableDeclaration" &&
        statement.declarations.some(
          (declaration) => bindingText(text, declaration.id) === "descriptor",
        ),
    );
    if (!descriptor)
      throw new TypeError("Range guard fixture lost its descriptor");
    expect(rawRangeInputs(text, program, [descriptor]).length > 0).toBe(
      index < bindings.length - 1,
    );
  }
});

test("every canonical range producer normalizes selection endpoints", async () => {
  const packageJson = await Bun.file(path.join(root, "package.json")).json();
  expect(packageJson.scripts["test:source-guards"].split(" ")).toContain(
    "scripts/canonical-range-producers.test.ts",
  );
  const found = new Map<string, string[]>();
  const glob = new Bun.Glob("**/*.ts");
  const directory = path.join(root, "packages/core/src/prosemirror");
  for await (const file of glob.scan({ cwd: directory })) {
    if (file.endsWith(".test.ts")) continue;
    const source = await Bun.file(path.join(directory, file)).text();
    const { program } = parseSourceModule(file, source);
    const roots: Node[] = [];
    walkAst(program, (node, ancestors) => {
      const lineage = [...ancestors, node];
      if (node.type === "ObjectExpression") {
        const type = typeOf(source, node);
        if (type !== undefined && Object.hasOwn(producers, type)) {
          roots.push(rangeInputRoot(source, lineage));
          const declaration = producerDeclaration(lineage);
          const owner = `${file}#${declarationName(source, declaration)}`;
          const entries = found.get(type) ?? [];
          entries.push(owner);
          found.set(type, entries);
          // Direct endpoints must be computed carrier positions; selection ranges use the helper.
          for (const property of node.properties) {
            const name = keyText(source, property);
            if (
              isPropertyAssignment(property) &&
              property.type === "Property" &&
              ["from", "to"].includes(name ?? "")
            ) {
              expect([
                "canonicalCommands.ts#canonicalRunFormatting",
                "extensions/marks/markUtils.ts#canonicalSetMarkFormatting",
                "extensions/marks/markUtils.ts#clearFormatting",
              ]).toContain(owner);
              expect(nodeText(source, property.value)).toBe(
                `representation.${name}`,
              );
            }
            if (property.type === "SpreadElement") {
              const range = nodeText(source, property.argument);
              if (range === "range") {
                expect(owner).toBe(
                  "extensions/marks/HyperlinkExtension.ts#removeHyperlink",
                );
              } else if (
                range.includes("Selection") ||
                range.includes("selection")
              ) {
                expect(range).toBe("canonicalSelectionRange(state)");
              }
            }
          }
          if (
            owner.endsWith("#canonicalSetMarkFormatting") ||
            owner.endsWith("#canonicalRunFormatting") ||
            owner.endsWith("#clearFormatting")
          ) {
            const declarationText =
              declaration && nodeText(source, declaration);
            expect(declarationText).toContain(
              "canonicalSelectionRanges(state)",
            );
            expect(declarationText).not.toMatch(/state\.selection\.(from|to)/u);
          } else {
            expect(
              node.properties.some(
                (property) =>
                  property.type === "SpreadElement" &&
                  (nodeText(source, property.argument) ===
                    "canonicalSelectionRange(state)" ||
                    (type === "removeHyperlink" &&
                      nodeText(source, property.argument) === "range")),
              ),
            ).toBe(true);
          }
        }
      }
      // Range intermediates cannot bypass normalization through destructured raw endpoints.
      if (
        node.type === "VariableDeclarator" &&
        node.init &&
        nodeText(source, node.init) === "state.selection" &&
        node.id.type === "ObjectPattern"
      ) {
        const names = node.id.properties.map((element) =>
          bindingKey(source, element),
        );
        const producer = declarationName(
          source,
          producerDeclaration(ancestors),
        );
        if (
          producer === "canonicalSetMarkFormatting" ||
          producer === "removalRange"
        ) {
          expect(names).not.toContain("from");
          expect(names).not.toContain("to");
        }
      }
    });
    if (roots.length > 0)
      expect(rawRangeInputs(source, program, roots)).toEqual([]);
    if (file === "extensions/marks/HyperlinkExtension.ts") {
      const removal = source.slice(
        source.indexOf("const removalRange"),
        source.indexOf("const removeHyperlink"),
      );
      expect(removal).toContain("canonicalSelectionRange(state)");
      expect(removal).not.toMatch(/state\.selection\.(from|to)/u);
    }
  }
  expect([...found.keys()].sort()).toEqual(Object.keys(producers).sort());
  for (const [type, expected] of Object.entries(producers)) {
    expect(found.get(type)?.sort()).toEqual([...expected].sort());
  }
});
