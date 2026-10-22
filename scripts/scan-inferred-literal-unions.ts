import { readdir } from "node:fs/promises";
import { join } from "node:path";

import ts from "typescript";

const root = join(import.meta.dir, "..");

const isLiteralUnion = (node: ts.TypeNode): boolean =>
  ts.isUnionTypeNode(node) &&
  node.types.length > 1 &&
  node.types.every(
    (member) =>
      ts.isLiteralTypeNode(member) &&
      (ts.isStringLiteral(member.literal) || ts.isNumericLiteral(member.literal)),
  );

const unwrapType = (node: ts.TypeNode): ts.TypeNode =>
  ts.isParenthesizedTypeNode(node) ? unwrapType(node.type) : node;

const hasInferredLiteralUnion = (node: ts.TypeNode): boolean => {
  const candidate = unwrapType(node);
  if (isLiteralUnion(candidate)) return true;
  if (ts.isArrayTypeNode(candidate)) return isLiteralUnion(unwrapType(candidate.elementType));
  if (ts.isTypeReferenceNode(candidate) && candidate.typeArguments?.length === 1) {
    return isLiteralUnion(unwrapType(candidate.typeArguments[0]!));
  }
  return false;
};

const declarationTypes = (source: ts.SourceFile): ReadonlyMap<string, ts.TypeNode> => {
  const result = new Map<string, ts.TypeNode>();
  for (const statement of source.statements) {
    if (ts.isVariableStatement(statement)) {
      for (const declaration of statement.declarationList.declarations) {
        if (ts.isIdentifier(declaration.name) && declaration.type) {
          result.set(declaration.name.text, declaration.type);
        }
      }
    } else if (ts.isFunctionDeclaration(statement) && statement.name && statement.type) {
      result.set(statement.name.text, statement.type);
    }
  }
  return result;
};

const unannotatedExports = (source: ts.SourceFile): ReadonlySet<string> => {
  const result = new Set<string>();
  for (const statement of source.statements) {
    const exported =
      ts.canHaveModifiers(statement) &&
      (ts.getModifiers(statement)?.some(({ kind }) => kind === ts.SyntaxKind.ExportKeyword) ??
        false);
    if (!exported) continue;
    if (ts.isVariableStatement(statement)) {
      for (const declaration of statement.declarationList.declarations) {
        if (ts.isIdentifier(declaration.name) && !declaration.type)
          result.add(declaration.name.text);
      }
    } else if (ts.isFunctionDeclaration(statement) && statement.name && !statement.type) {
      result.add(statement.name.text);
    }
  }
  return result;
};

const findings: string[] = [];
const packageEntries = await readdir(join(root, "packages"), { withFileTypes: true });
for (const entry of packageEntries) {
  if (!entry.isDirectory()) continue;
  const packageDirectory = join(root, "packages", entry.name);
  const manifest = Bun.file(join(packageDirectory, "package.json"));
  if (!(await manifest.exists())) continue;
  if (((await manifest.json()) as { private?: boolean }).private) continue;
  const distDirectory = join(packageDirectory, "dist");
  let declarationCount = 0;
  const glob = new Bun.Glob("**/*.d.ts");
  for await (const relativeFile of glob.scan({ cwd: distDirectory, onlyFiles: true })) {
    declarationCount += 1;
    const declarationText = await Bun.file(join(distDirectory, relativeFile)).text();
    const declarationSource = ts.createSourceFile(
      relativeFile,
      declarationText,
      ts.ScriptTarget.Latest,
      true,
    );
    const emittedTypes = declarationTypes(declarationSource);
    if (![...emittedTypes.values()].some(hasInferredLiteralUnion)) continue;

    const sourceRelative = relativeFile.replace(/\.d\.ts$/u, ".ts");
    const sourceCandidates = [
      join(packageDirectory, "src", sourceRelative),
      join(packageDirectory, "src", sourceRelative.replace(/\.ts$/u, ".tsx")),
    ];
    const sourceFile = sourceCandidates.find((candidate) => Bun.file(candidate).size > 0);
    if (!sourceFile) continue;
    const sourceText = await Bun.file(sourceFile).text();
    const source = ts.createSourceFile(sourceFile, sourceText, ts.ScriptTarget.Latest, true);
    const unannotated = unannotatedExports(source);
    for (const [name, type] of emittedTypes) {
      if (unannotated.has(name) && hasInferredLiteralUnion(type)) {
        findings.push(
          `${sourceFile.slice(root.length + 1)}:${name} -> ${type.getText(declarationSource)}`,
        );
      }
    }
  }
  if (declarationCount === 0) {
    throw new Error(`${entry.name}: dist has no declaration files; build published packages first`);
  }
}

if (findings.length > 0) {
  console.error("Inferred literal unions in published declarations:\n" + findings.join("\n"));
  process.exit(1);
}
console.log("No inferred string or number literal unions found in published declarations.");
