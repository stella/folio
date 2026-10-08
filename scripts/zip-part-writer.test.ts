/**
 * ZIP package entries have one write boundary. JSZip's `file()` and `folder()`
 * create parent directory entries by default; that changes a DOCX package's
 * entry set even when the caller only meant to replace one part. Keep every
 * production writer behind `writeZipPart`, which fixes both folder creation
 * and the ZIP date used for reproducible output.
 */

import { describe, expect, test } from "bun:test";
import ts from "typescript";
import path from "node:path";

const REPOSITORY_ROOT = path.resolve(import.meta.dir, "..");
const PACKAGES_DIR = path.join(REPOSITORY_ROOT, "packages");
const WRITER = "docx-core/src/zip/writeZipPart.ts";

type Violation = { line: number; method: "file" | "folder" };

const staticPropertyName = (expression: ts.Expression): "file" | "folder" | undefined => {
  if (ts.isPropertyAccessExpression(expression)) {
    return expression.name.text === "file" || expression.name.text === "folder"
      ? expression.name.text
      : undefined;
  }
  if (!ts.isElementAccessExpression(expression)) return undefined;
  const argument = expression.argumentExpression;
  if (
    argument !== undefined &&
    (ts.isStringLiteral(argument) || ts.isNoSubstitutionTemplateLiteral(argument)) &&
    (argument.text === "file" || argument.text === "folder")
  ) {
    return argument.text;
  }
  return undefined;
};

const unwrapExpression = (expression: ts.Expression): ts.Expression => {
  if (
    ts.isParenthesizedExpression(expression) ||
    ts.isAsExpression(expression) ||
    ts.isTypeAssertionExpression(expression) ||
    ts.isNonNullExpression(expression) ||
    ts.isAwaitExpression(expression) ||
    ts.isSatisfiesExpression(expression)
  ) {
    return unwrapExpression(expression.expression);
  }
  return expression;
};

const receiverExpression = (expression: ts.Expression): ts.Expression | undefined => {
  const unwrapped = unwrapExpression(expression);
  if (ts.isPropertyAccessExpression(unwrapped)) return unwrapped.expression;
  if (ts.isElementAccessExpression(unwrapped)) return unwrapped.expression;
  return undefined;
};

const findWrites = (sourceText: string, fileName = "fixture.ts"): Violation[] => {
  const source = ts.createSourceFile(
    fileName,
    sourceText,
    ts.ScriptTarget.Latest,
    true,
    fileName.endsWith(".tsx") ? ts.ScriptKind.TSX : ts.ScriptKind.TS,
  );
  const violations: Violation[] = [];
  const jszipBindings = new Set<string>();
  const declarations: ts.VariableDeclaration[] = [];
  const typedDeclarations: Array<ts.ParameterDeclaration | ts.VariableDeclaration> = [];
  const typedZipReceivers = new Set<string>();
  const directFileReceivers = new Set<string>();

  const collectJsZipBindings = (node: ts.Node): void => {
    if (
      ts.isImportDeclaration(node) &&
      ts.isStringLiteral(node.moduleSpecifier) &&
      node.moduleSpecifier.text === "jszip" &&
      node.importClause !== undefined
    ) {
      if (node.importClause.name !== undefined) jszipBindings.add(node.importClause.name.text);
      const bindings = node.importClause.namedBindings;
      if (bindings !== undefined && ts.isNamespaceImport(bindings)) {
        jszipBindings.add(bindings.name.text);
      } else if (bindings !== undefined && ts.isNamedImports(bindings)) {
        for (const element of bindings.elements) {
          if (element.propertyName?.text === "JSZip" || element.name.text === "JSZip") {
            jszipBindings.add(element.name.text);
          }
        }
      }
    }
    if (ts.isVariableDeclaration(node)) declarations.push(node);
    if (ts.isParameter(node) || ts.isVariableDeclaration(node)) typedDeclarations.push(node);
  };
  const allNodes: ts.Node[] = [source];
  while (allNodes.length > 0) {
    const node = allNodes.pop();
    if (node === undefined) continue;
    collectJsZipBindings(node);
    allNodes.push(...node.getChildren(source));
  }
  for (const declaration of typedDeclarations) {
    if (declaration.type === undefined || !ts.isTypeReferenceNode(declaration.type)) continue;
    if (!jszipBindings.has(declaration.type.typeName.getText(source))) continue;
    typedZipReceivers.add(declaration.name.getText(source));
  }

  const isJSZipConstruction = (initializer: ts.Expression): boolean => {
    const expression = unwrapExpression(initializer);
    if (
      ts.isNewExpression(expression) &&
      ts.isIdentifier(expression.expression) &&
      jszipBindings.has(expression.expression.text)
    ) {
      return true;
    }
    if (!ts.isCallExpression(expression) || !ts.isPropertyAccessExpression(expression.expression)) {
      return false;
    }
    const callee = expression.expression;
    return (
      callee.name.text === "loadAsync" &&
      ts.isIdentifier(callee.expression) &&
      jszipBindings.has(callee.expression.text)
    );
  };

  const knownZipReceivers = new Set(typedZipReceivers);
  let addedReceiver = true;
  while (addedReceiver) {
    addedReceiver = false;
    for (const declaration of declarations) {
      if (!ts.isIdentifier(declaration.name) || declaration.initializer === undefined) continue;
      const initializer = unwrapExpression(declaration.initializer);
      const isZipAlias = ts.isIdentifier(initializer) && knownZipReceivers.has(initializer.text);
      if (
        (isZipAlias || isJSZipConstruction(initializer)) &&
        !knownZipReceivers.has(declaration.name.text)
      ) {
        knownZipReceivers.add(declaration.name.text);
        addedReceiver = true;
      }
    }
  }

  const receiverKey = (expression: ts.Expression): string =>
    unwrapExpression(expression).getText(source);
  const collectDirectFileReceivers = (node: ts.Node): void => {
    if (!ts.isCallExpression(node) || staticPropertyName(node.expression) !== "file") return;
    const receiver = receiverExpression(node.expression);
    if (receiver !== undefined) directFileReceivers.add(receiverKey(receiver));
  };
  const collectReceivers: ts.Node[] = [source];
  while (collectReceivers.length > 0) {
    const node = collectReceivers.pop();
    if (node === undefined) continue;
    collectDirectFileReceivers(node);
    collectReceivers.push(...node.getChildren(source));
  }

  const addViolation = (node: ts.Node, method: "file" | "folder"): void => {
    const { line } = source.getLineAndCharacterOfPosition(node.getStart(source));
    violations.push({ line: line + 1, method });
  };
  const inspect = (node: ts.Node): void => {
    if (ts.isCallExpression(node)) {
      const method = staticPropertyName(node.expression);
      if (
        method !== undefined &&
        (method === "folder" ||
          node.arguments.some(ts.isSpreadElement) ||
          (method === "file" && node.arguments.length >= 2))
      ) {
        addViolation(node, method);
      }
    }

    if (ts.isPropertyAccessExpression(node) || ts.isElementAccessExpression(node)) {
      const method = staticPropertyName(node);
      if (method === undefined) return;
      const directCall = ts.isCallExpression(node.parent) && node.parent.expression === node;
      const receiver = receiverExpression(node);
      const receiverId = receiver === undefined ? undefined : receiverKey(receiver);
      const isKnownZipReceiver =
        receiverId !== undefined &&
        (knownZipReceivers.has(receiverId) || directFileReceivers.has(receiverId));
      if (!directCall && isKnownZipReceiver) addViolation(node, method);
    }

    if (ts.isBindingElement(node)) {
      const propertyName = node.propertyName;
      let name: string | undefined;
      if (propertyName === undefined && ts.isIdentifier(node.name)) name = node.name.text;
      if (
        propertyName !== undefined &&
        (ts.isIdentifier(propertyName) || ts.isStringLiteral(propertyName))
      )
        name = propertyName.text;
      if (name === "file" || name === "folder") {
        const declaration = node.parent.parent;
        const receiver =
          ts.isVariableDeclaration(declaration) && declaration.initializer !== undefined
            ? receiverKey(declaration.initializer)
            : undefined;
        if (receiver !== undefined && knownZipReceivers.has(receiver)) addViolation(node, name);
      }
    }
  };
  const visit = (node: ts.Node): void => {
    inspect(node);
    ts.forEachChild(node, visit);
  };
  visit(source);
  return violations;
};

const isExcludedSource = (relativePath: string): boolean => {
  const segments = relativePath.split(path.sep);
  return (
    segments.some((segment) => /^(?:__tests__|__fixtures__|tests?|fixtures?)$/u.test(segment)) ||
    /\.(?:test|spec|fixture)\.tsx?$/u.test(relativePath)
  );
};

const isGuardedSource = (relativePath: string): boolean =>
  relativePath !== WRITER && !isExcludedSource(relativePath);

const sourceViolations = (relativePath: string, sourceText: string): Violation[] =>
  isGuardedSource(relativePath) ? findWrites(sourceText, relativePath) : [];

const productionViolations = async (): Promise<string[]> => {
  const offenders: string[] = [];
  for await (const file of new Bun.Glob("*/src/**/*.{ts,tsx}").scan({
    absolute: true,
    cwd: PACKAGES_DIR,
  })) {
    const relative = path.relative(PACKAGES_DIR, file);
    if (!isGuardedSource(relative)) continue;
    const sourceText = await Bun.file(file).text();
    for (const violation of sourceViolations(relative, sourceText)) {
      offenders.push(`${relative}:${violation.line}  raw JSZip ${violation.method} write`);
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

  test("production ZIP writes use the owner helper", async () => {
    expect(await productionViolations()).toEqual([]);
  });
});
