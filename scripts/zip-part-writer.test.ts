import { describe, expect, setDefaultTimeout, test } from "bun:test";
/**
 * ZIP package entries have one write boundary. JSZip's `file()` and `folder()`
 * create parent directory entries by default; that changes a DOCX package's
 * entry set even when the caller only meant to replace one part. Keep every
 * production writer behind `writeZipPart`, which fixes both folder creation
 * and the ZIP date used for reproducible output.
 */
import path from "node:path";
import type { VariableDeclarator } from "oxc-parser";

import { scanTestTimeout } from "../test/scan-timeout";
import {
  bindingElements,
  collectSyntaxTree,
  createLineIndex,
  type DeclaredName,
  isFunctionLike,
  type Node,
  nodeText,
  parameterParts,
  parametersOf,
  parseLikeTypeScript,
} from "./lib/oxc-ast";

setDefaultTimeout(scanTestTimeout(30_000));

const REPOSITORY_ROOT = path.resolve(import.meta.dir, "..");
const PACKAGES_DIR = path.join(REPOSITORY_ROOT, "packages");
const WRITER = "utf8-boundaries/src/zip.ts";

type Violation = { line: number; method: "file" | "folder" };

const isFileOrFolder = (name: string | undefined): name is "file" | "folder" =>
  name === "file" || name === "folder";

const staticPropertyName = (
  expression: Node,
): "file" | "folder" | undefined => {
  if (expression.type !== "MemberExpression") return undefined;
  if (!expression.computed) {
    const name =
      expression.property.type === "Identifier"
        ? expression.property.name
        : undefined;
    return isFileOrFolder(name) ? name : undefined;
  }
  const argument = expression.property;
  let text: string | undefined;
  if (argument.type === "Literal" && typeof argument.value === "string")
    text = argument.value;
  if (argument.type === "TemplateLiteral" && argument.expressions.length === 0)
    text = argument.quasis[0]?.value.cooked ?? undefined;
  return isFileOrFolder(text) ? text : undefined;
};

const unwrapExpression = (expression: Node): Node => {
  if (
    expression.type === "ParenthesizedExpression" ||
    expression.type === "TSAsExpression" ||
    expression.type === "TSTypeAssertion" ||
    expression.type === "TSNonNullExpression" ||
    expression.type === "TSSatisfiesExpression"
  ) {
    return unwrapExpression(expression.expression);
  }
  if (expression.type === "AwaitExpression")
    return unwrapExpression(expression.argument);
  return expression;
};

const receiverExpression = (expression: Node): Node | undefined => {
  const unwrapped = unwrapExpression(expression);
  if (unwrapped.type === "MemberExpression") return unwrapped.object;
  return undefined;
};

const findWrites = (
  sourceText: string,
  fileName = "fixture.ts",
): Violation[] => {
  const { program } = parseLikeTypeScript(
    fileName,
    sourceText,
    /\.[jt]sx$/u.test(fileName) ? "tsx" : "ts",
  );
  const { nodes, parentOf } = collectSyntaxTree(program);
  const lineOf = createLineIndex(sourceText);
  const violations: Violation[] = [];
  const jszipBindings = new Set<string>();
  const declarations: VariableDeclarator[] = [];
  const typedDeclarations: DeclaredName[] = [];
  const typedZipReceivers = new Set<string>();
  const directFileReceivers = new Set<string>();

  for (const node of nodes) {
    if (
      node.type === "ImportDeclaration" &&
      node.source.value === "jszip" &&
      node.specifiers.length > 0
    ) {
      for (const specifier of node.specifiers) {
        if (specifier.type !== "ImportSpecifier") {
          jszipBindings.add(specifier.local.name);
          continue;
        }
        const imported =
          specifier.imported.type === "Literal"
            ? specifier.imported.value
            : specifier.imported.name;
        if (imported === "JSZip" || specifier.local.name === "JSZip")
          jszipBindings.add(specifier.local.name);
      }
    }
    if (node.type === "VariableDeclarator") {
      declarations.push(node);
      // A declarator's type sits on its binding name, as a parameter's does.
      typedDeclarations.push(parameterParts(node.id));
      if (
        node.id.type === "Identifier" &&
        node.init?.type === "CallExpression" &&
        node.init.callee.type === "Identifier" &&
        node.init.callee.name === "require" &&
        node.init.arguments.some(
          (argument) =>
            argument.type === "Literal" && argument.value === "jszip",
        )
      ) {
        jszipBindings.add(node.id.name);
      }
    }
    if (isFunctionLike(node))
      for (const parameter of parametersOf(node))
        typedDeclarations.push(parameterParts(parameter));
  }
  // TypeScript's text of a binding name stops before its annotation.
  const bindingNameText = ({ name, type }: DeclaredName): string =>
    name.type === "Identifier"
      ? name.name
      : sourceText
          .slice(name.start, sourceText.lastIndexOf(":", type?.start))
          .trimEnd();
  for (const declaration of typedDeclarations) {
    const { type } = declaration;
    if (type === undefined || type.type !== "TSTypeReference") continue;
    if (!jszipBindings.has(nodeText(sourceText, type.typeName))) continue;
    typedZipReceivers.add(bindingNameText(declaration));
  }

  const isJSZipConstruction = (initializer: Node): boolean => {
    const expression = unwrapExpression(initializer);
    if (
      expression.type === "NewExpression" &&
      expression.callee.type === "Identifier" &&
      jszipBindings.has(expression.callee.name)
    ) {
      return true;
    }
    if (
      expression.type !== "CallExpression" ||
      expression.callee.type !== "MemberExpression" ||
      expression.callee.computed
    ) {
      return false;
    }
    const callee = expression.callee;
    return (
      callee.property.type === "Identifier" &&
      callee.property.name === "loadAsync" &&
      callee.object.type === "Identifier" &&
      jszipBindings.has(callee.object.name)
    );
  };

  const knownZipReceivers = new Set(typedZipReceivers);
  let addedReceiver = true;
  while (addedReceiver) {
    addedReceiver = false;
    for (const declaration of declarations) {
      if (declaration.id.type !== "Identifier" || !declaration.init) continue;
      const initializer = unwrapExpression(declaration.init);
      const isZipAlias =
        initializer.type === "Identifier" &&
        knownZipReceivers.has(initializer.name);
      if (
        (isZipAlias || isJSZipConstruction(initializer)) &&
        !knownZipReceivers.has(declaration.id.name)
      ) {
        knownZipReceivers.add(declaration.id.name);
        addedReceiver = true;
      }
    }
  }

  const receiverKey = (expression: Node): string =>
    nodeText(sourceText, unwrapExpression(expression));
  for (const node of nodes) {
    if (
      node.type !== "CallExpression" ||
      staticPropertyName(node.callee) !== "file"
    )
      continue;
    const receiver = receiverExpression(node.callee);
    if (receiver !== undefined) directFileReceivers.add(receiverKey(receiver));
  }

  const addViolation = (node: Node, method: "file" | "folder"): void => {
    const { line } = lineOf(node.start);
    violations.push({ line: line + 1, method });
  };
  for (const node of nodes) {
    if (node.type === "CallExpression") {
      const method = staticPropertyName(node.callee);
      if (
        method !== undefined &&
        (method === "folder" ||
          node.arguments.some(
            (argument) => argument.type === "SpreadElement",
          ) ||
          (method === "file" && node.arguments.length >= 2))
      ) {
        addViolation(node, method);
      }
    }

    if (node.type === "MemberExpression") {
      const method = staticPropertyName(node);
      if (method === undefined) continue;
      const parent = parentOf(node);
      const directCall =
        parent?.type === "CallExpression" && parent.callee === node;
      const receiver = receiverExpression(node);
      const receiverId =
        receiver === undefined ? undefined : receiverKey(receiver);
      const isKnownZipReceiver =
        receiverId !== undefined &&
        (knownZipReceivers.has(receiverId) ||
          directFileReceivers.has(receiverId));
      if (!directCall && isKnownZipReceiver) addViolation(node, method);
    }

    if (node.type === "ObjectPattern" || node.type === "ArrayPattern") {
      const declaration = parentOf(node);
      const receiver =
        declaration?.type === "VariableDeclarator" && declaration.init
          ? receiverKey(declaration.init)
          : undefined;
      for (const element of bindingElements(node)) {
        const propertyName = element.propertyName;
        let name: string | undefined;
        if (propertyName === undefined && element.name.type === "Identifier")
          name = element.name.name;
        if (
          propertyName !== undefined &&
          !(element.node.type === "Property" && element.node.computed)
        ) {
          if (propertyName.type === "Identifier") name = propertyName.name;
          if (
            propertyName.type === "Literal" &&
            typeof propertyName.value === "string"
          )
            name = propertyName.value;
        }
        if (
          isFileOrFolder(name) &&
          receiver !== undefined &&
          knownZipReceivers.has(receiver)
        )
          addViolation(element.node, name);
      }
    }
  }
  return violations;
};

const isExcludedSource = (relativePath: string): boolean => {
  const segments = relativePath.split(path.sep);
  return (
    segments.some((segment) =>
      /^(?:__tests__|__fixtures__|tests?|fixtures?)$/u.test(segment),
    ) || /\.(?:test|spec|fixture)\.[cm]?[jt]sx?$/u.test(relativePath)
  );
};

const isGuardedSource = (relativePath: string): boolean =>
  relativePath !== WRITER && !isExcludedSource(relativePath);

const sourceViolations = (
  relativePath: string,
  sourceText: string,
): Violation[] =>
  isGuardedSource(relativePath) ? findWrites(sourceText, relativePath) : [];

const productionViolations = async (): Promise<string[]> => {
  const offenders: string[] = [];
  for await (const file of new Bun.Glob(
    "*/src/**/*.{ts,tsx,js,jsx,mts,cts,mjs,cjs}",
  ).scan({
    absolute: true,
    cwd: PACKAGES_DIR,
  })) {
    const relative = path.relative(PACKAGES_DIR, file);
    if (!isGuardedSource(relative)) continue;
    const sourceText = await Bun.file(file).text();
    for (const violation of sourceViolations(relative, sourceText)) {
      offenders.push(
        `${relative}:${violation.line}  raw JSZip ${violation.method} write`,
      );
    }
  }
  return offenders.sort();
};

describe("ZIP part writer", () => {
  test("the fixture detector rejects writes and escaping writer references", () => {
    const fixture = `
      import JSZip from "jszip";
      const zip = new JSZip();
      zip.file("word/document.xml", xml);
      zip["file"]("word/document.xml", xml);
      zip?.file?.("word/document.xml", xml);
      zip.folder("word");
      const { file: addPart } = zip;
      addPart("word/document.xml", xml);
      zip.file(...args);
      zip.file.bind(zip)("word/x", xml);
      const bound = zip.file.bind(zip);
      bound("word/x", xml);
    `;
    const violations = findWrites(fixture);
    expect(violations).toHaveLength(8);
    expect(violations.some(({ method }) => method === "folder")).toBe(true);
  });

  test("reads and the owner helper are permitted", () => {
    expect(
      findWrites(`
      const source = zip.file("word/document.xml");
      const localFile = Bun.file("input.docx");
      const contents = zip["file"]("word/styles.xml")?.async("text");
      const readMethod = (path: string) => zip.file(path);
    `),
    ).toEqual([]);
    expect(sourceViolations(WRITER, `zip.file(name, content);`)).toEqual([]);
    expect(
      findWrites(`
        const fieldIds = { file: "file-id" };
        const built = { file: "built-file" };
        fieldIds.file;
        built.file;
      `),
    ).toEqual([]);
  });

  test("every package rejects direct string inputs before archive generation", async () => {
    for await (const manifest of new Bun.Glob("*/package.json").scan({
      cwd: PACKAGES_DIR,
    })) {
      const unsafeSource = `
        import JSZip from "jszip";
        const zip = new JSZip();
        zip.file("word/document.xml", "a".repeat(16383) + "👩");
        zip.generateAsync({ type: "uint8array" });
      `;
      for (const extension of [
        "ts",
        "tsx",
        "js",
        "jsx",
        "mts",
        "cts",
        "mjs",
        "cjs",
      ]) {
        const sourcePath = path.join(
          path.dirname(manifest),
          "src",
          `unsafe.${extension}`,
        );
        expect(sourceViolations(sourcePath, unsafeSource)).toHaveLength(1);
      }
    }
  });

  test("CommonJS writers cannot escape through bound method references", () => {
    expect(
      findWrites(`
        const Archive = require("jszip");
        const zip = new Archive();
        const write = zip.file.bind(zip);
        write("word/document.xml", "text");
        zip.generateAsync({ type: "uint8array" });
      `),
    ).toHaveLength(1);
  });

  test(
    "production ZIP writes use the owner helper",
    async () => {
      expect(await productionViolations()).toEqual([]);
    },
    scanTestTimeout(30_000),
  );
});
