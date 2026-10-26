import { readdir } from "node:fs/promises";
import { join } from "node:path";
import type { TSType } from "oxc-parser";

import { nodeText, parseSourceModule, type Program } from "./lib/oxc-ast";

const root = join(import.meta.dir, "..");

const isLiteralUnion = (node: TSType): boolean =>
  node.type === "TSUnionType" &&
  node.types.length > 1 &&
  node.types.every((member) => {
    if (member.type !== "TSLiteralType" || member.literal.type !== "Literal")
      return false;
    const { value } = member.literal;
    return typeof value === "string" || typeof value === "number";
  });

const unwrapType = (node: TSType): TSType =>
  node.type === "TSParenthesizedType" ? unwrapType(node.typeAnnotation) : node;

const hasInferredLiteralUnion = (node: TSType): boolean => {
  const candidate = unwrapType(node);
  if (isLiteralUnion(candidate)) return true;
  if (candidate.type === "TSArrayType")
    return isLiteralUnion(unwrapType(candidate.elementType));
  const typeArgument =
    candidate.type === "TSTypeReference" &&
    candidate.typeArguments?.params.length === 1
      ? candidate.typeArguments.params.at(0)
      : undefined;
  return typeArgument !== undefined && isLiteralUnion(unwrapType(typeArgument));
};

type TopLevelDeclaration = {
  readonly name: string;
  readonly type: TSType | undefined;
  readonly exported: boolean;
};

/** Top-level variables and functions, including exported and declared ones. */
const topLevelDeclarations = (program: Program): TopLevelDeclaration[] =>
  program.body.flatMap((statement): TopLevelDeclaration[] => {
    const declaration =
      statement.type === "ExportNamedDeclaration" ||
      statement.type === "ExportDefaultDeclaration"
        ? statement.declaration
        : statement;
    const exported = declaration !== statement;
    if (declaration?.type === "VariableDeclaration")
      return declaration.declarations.flatMap(({ id }) =>
        id.type === "Identifier"
          ? [
              {
                name: id.name,
                type: id.typeAnnotation?.typeAnnotation,
                exported,
              },
            ]
          : [],
      );
    if (
      (declaration?.type === "FunctionDeclaration" ||
        declaration?.type === "TSDeclareFunction") &&
      declaration.id !== null
    )
      return [
        {
          name: declaration.id.name,
          type: declaration.returnType?.typeAnnotation,
          exported,
        },
      ];
    return [];
  });

const declarationTypes = (program: Program): ReadonlyMap<string, TSType> => {
  const result = new Map<string, TSType>();
  for (const { name, type } of topLevelDeclarations(program)) {
    if (type !== undefined) result.set(name, type);
  }
  return result;
};

const unannotatedExports = (program: Program): ReadonlySet<string> =>
  new Set(
    topLevelDeclarations(program)
      .filter(({ type, exported }) => exported && type === undefined)
      .map(({ name }) => name),
  );

const findings: string[] = [];
const packageEntries = await readdir(join(root, "packages"), {
  withFileTypes: true,
});
for (const entry of packageEntries) {
  if (!entry.isDirectory()) continue;
  const packageDirectory = join(root, "packages", entry.name);
  const manifest = Bun.file(join(packageDirectory, "package.json"));
  if (!(await manifest.exists())) continue;
  if (((await manifest.json()) as { private?: boolean }).private) continue;
  const distDirectory = join(packageDirectory, "dist");
  let declarationCount = 0;
  const glob = new Bun.Glob("**/*.d.ts");
  for await (const relativeFile of glob.scan({
    cwd: distDirectory,
    onlyFiles: true,
  })) {
    declarationCount += 1;
    const declarationText = await Bun.file(
      join(distDirectory, relativeFile),
    ).text();
    const emittedTypes = declarationTypes(
      parseSourceModule(relativeFile, declarationText).program,
    );
    if (![...emittedTypes.values()].some(hasInferredLiteralUnion)) continue;

    const sourceRelative = relativeFile.replace(/\.d\.ts$/u, ".ts");
    const sourceCandidates = [
      join(packageDirectory, "src", sourceRelative),
      join(packageDirectory, "src", sourceRelative.replace(/\.ts$/u, ".tsx")),
    ];
    const sourceFile = sourceCandidates.find(
      (candidate) => Bun.file(candidate).size > 0,
    );
    if (!sourceFile) continue;
    const sourceText = await Bun.file(sourceFile).text();
    const unannotated = unannotatedExports(
      parseSourceModule(sourceFile, sourceText).program,
    );
    for (const [name, type] of emittedTypes) {
      if (unannotated.has(name) && hasInferredLiteralUnion(type)) {
        findings.push(
          `${sourceFile.slice(root.length + 1)}:${name} -> ${nodeText(declarationText, type)}`,
        );
      }
    }
  }
  if (declarationCount === 0) {
    throw new Error(
      `${entry.name}: dist has no declaration files; build published packages first`,
    );
  }
}

if (findings.length > 0) {
  console.error(
    "Inferred literal unions in published declarations:\n" +
      findings.join("\n"),
  );
  process.exit(1);
}
console.log(
  "No inferred string or number literal unions found in published declarations.",
);
