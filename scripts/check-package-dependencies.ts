#!/usr/bin/env bun

import { existsSync, readFileSync } from "node:fs";
import { builtinModules } from "node:module";
import path from "node:path";
import ts from "typescript";

import { PUBLISHED_PACKAGES } from "./lib/published-packages";

export type DependencyManifest = {
  name: string;
  dependencies?: Record<string, string>;
  optionalDependencies?: Record<string, string>;
  devDependencies?: Record<string, string>;
  peerDependencies?: Record<string, string>;
  peerDependenciesMeta?: Record<string, { optional?: boolean }>;
};

type PackageImport = { specifier: string; typeOnly: boolean };
const BUILTINS = new Set(
  builtinModules.flatMap((name) => [name, `node:${name}`]),
);

export const importedPackages = (
  source: string,
  fileName: string,
): PackageImport[] => {
  const imports: PackageImport[] = [];
  if (fileName.endsWith(".css")) {
    const content = source.replace(/\/\*[\s\S]*?\*\//gu, "");
    for (const match of content.matchAll(
      /@import\s+(?:url\(\s*)?["']([^"']+)["']/gu,
    )) {
      const specifier = match.at(1);
      if (specifier) imports.push({ specifier, typeOnly: false });
    }
    return imports;
  }
  const content = fileName.endsWith(".vue")
    ? [...source.matchAll(/<script\b[^>]*>([\s\S]*?)<\/script>/gu)]
        .map((match) => match.at(1) ?? "")
        .join("\n")
    : source;
  const scriptKind = /\.[jt]sx$/u.test(fileName)
    ? ts.ScriptKind.TSX
    : ts.ScriptKind.TS;
  const file = ts.createSourceFile(
    fileName,
    content,
    ts.ScriptTarget.Latest,
    true,
    scriptKind,
  );
  const add = (node: ts.Node, typeOnly: boolean) => {
    if (ts.isStringLiteralLike(node))
      imports.push({ specifier: node.text, typeOnly });
  };
  const visit = (node: ts.Node) => {
    if (ts.isImportDeclaration(node)) {
      const clause = node.importClause;
      const bindings = clause?.namedBindings;
      const typeOnly =
        clause?.isTypeOnly === true ||
        (clause !== undefined &&
          clause.name === undefined &&
          bindings !== undefined &&
          ts.isNamedImports(bindings) &&
          bindings.elements.length > 0 &&
          bindings.elements.every((element) => element.isTypeOnly));
      add(node.moduleSpecifier, typeOnly);
    } else if (ts.isExportDeclaration(node) && node.moduleSpecifier) {
      const typeOnly =
        node.isTypeOnly ||
        (node.exportClause !== undefined &&
          ts.isNamedExports(node.exportClause) &&
          node.exportClause.elements.length > 0 &&
          node.exportClause.elements.every((element) => element.isTypeOnly));
      add(node.moduleSpecifier, typeOnly);
    } else if (
      ts.isImportEqualsDeclaration(node) &&
      ts.isExternalModuleReference(node.moduleReference)
    ) {
      if (node.moduleReference.expression)
        add(node.moduleReference.expression, node.isTypeOnly);
    } else if (
      ts.isImportTypeNode(node) &&
      ts.isLiteralTypeNode(node.argument)
    ) {
      add(node.argument.literal, true);
    } else if (
      ts.isCallExpression(node) &&
      (node.expression.kind === ts.SyntaxKind.ImportKeyword ||
        (ts.isIdentifier(node.expression) &&
          node.expression.text === "require"))
    ) {
      const argument = node.arguments.at(0);
      if (argument) add(argument, false);
    }
    ts.forEachChild(node, visit);
  };
  visit(file);
  for (const reference of file.typeReferenceDirectives) {
    imports.push({ specifier: reference.fileName, typeOnly: true });
  }
  return imports;
};

const packageName = (specifier: string) =>
  specifier
    .split("/")
    .slice(0, specifier.startsWith("@") ? 2 : 1)
    .join("/");

type CheckImportsOptions = {
  manifest: DependencyManifest;
  source: string;
  fileName: string;
};

export const checkImports = ({
  manifest,
  source,
  fileName,
}: CheckImportsOptions): string[] => {
  const runtime = new Set([
    manifest.name,
    ...Object.keys(manifest.dependencies ?? {}),
    ...Object.keys(manifest.optionalDependencies ?? {}),
    ...Object.keys(manifest.peerDependencies ?? {}),
  ]);
  const development = new Set([
    ...runtime,
    ...Object.keys(manifest.devDependencies ?? {}),
  ]);
  const declaration = /\.d\.[cm]?ts$/u.test(fileName);
  return importedPackages(source, fileName).flatMap(
    ({ specifier, typeOnly }) => {
      if (
        specifier.startsWith(".") ||
        specifier.startsWith("/") ||
        BUILTINS.has(specifier) ||
        /^(?:https?:|data:)/u.test(specifier)
      )
        return [];
      const name = packageName(specifier);
      // Source declaration references and stylesheet preprocessors are build inputs.
      // Emitted declarations and CSS must resolve from the consumer manifest.
      const sourceBuildInput =
        fileName.startsWith("src/") && (typeOnly || fileName.endsWith(".css"));
      const declared = sourceBuildInput ? development : runtime;
      if (declared.has(name)) return [];
      // TypeScript resolves ambient type packages through @types when a package
      // does not ship declarations. The declaration must still own that dependency.
      const typesName = `@types/${name.startsWith("@") ? name.slice(1).replace("/", "__") : name}`;
      if (typeOnly && declared.has(typesName)) return [];
      return [
        `${manifest.name}: ${fileName} imports undeclared ${specifier}${declaration ? " in consumer declarations" : ""}`,
      ];
    },
  );
};

export const checkRequiredPeers = (
  manifest: DependencyManifest,
  dependency: DependencyManifest,
): string[] =>
  Object.keys(dependency.peerDependencies ?? {}).flatMap((peer) => {
    if (dependency.peerDependenciesMeta?.[peer]?.optional === true) return [];
    if (manifest.dependencies?.[peer] !== undefined) return [];
    // An optional parent peer lets consumers omit what the dependency requires.
    if (
      manifest.peerDependencies?.[peer] !== undefined &&
      manifest.peerDependenciesMeta?.[peer]?.optional !== true
    )
      return [];
    const reason =
      manifest.peerDependencies?.[peer] === undefined
        ? "undeclared"
        : "optional";
    return [
      `${manifest.name}: dependency ${dependency.name} requires ${reason} peer ${peer}`,
    ];
  });

const readManifest = (file: string): DependencyManifest =>
  JSON.parse(readFileSync(file, "utf8"));

const installedManifest = (root: string, name: string) => {
  const workspace = PUBLISHED_PACKAGES.find((pkg) => pkg.name === name);
  if (workspace) return readManifest(path.join(workspace.root, "package.json"));
  let directory = root;
  for (;;) {
    const manifest = path.join(directory, "node_modules", name, "package.json");
    if (existsSync(manifest)) return readManifest(manifest);
    const parent = path.dirname(directory);
    if (parent === directory) return undefined;
    directory = parent;
  }
};

// Tests and fixtures are not production modules. The builds and package-content
// gate exclude these paths from the emitted consumer surface.
const isProduction = (file: string) =>
  !/(?:^|\/)(?:__tests__|__fixtures__|fixtures|__snapshots__)(?:\/|$)|\.(?:test|spec)\.[^/]+$/u.test(
    file,
  );

type CheckPublishedDependenciesOptions = { emitted?: boolean };

export const checkPublishedDependencies = async ({
  emitted = false,
}: CheckPublishedDependenciesOptions = {}) => {
  const issues: string[] = [];
  for (const pkg of PUBLISHED_PACKAGES) {
    const manifest = readManifest(path.join(pkg.root, "package.json"));
    for (const directory of emitted ? ["dist"] : ["src"]) {
      const root = path.join(pkg.root, directory);
      if (!existsSync(root)) {
        issues.push(`${pkg.name}: missing ${directory} directory`);
        continue;
      }
      for await (const file of new Bun.Glob(
        "**/*.{ts,tsx,mts,cts,js,jsx,mjs,cjs,vue,css}",
      ).scan(root)) {
        if (!isProduction(file)) continue;
        issues.push(
          ...checkImports({
            manifest,
            source: readFileSync(path.join(root, file), "utf8"),
            fileName: `${directory}/${file}`,
          }),
        );
      }
    }
    for (const dependency of new Set([
      ...Object.keys(manifest.dependencies ?? {}),
      ...Object.keys(manifest.optionalDependencies ?? {}),
    ])) {
      const installed = installedManifest(pkg.root, dependency);
      if (!installed) {
        issues.push(
          `${pkg.name}: dependency ${dependency} is not installed; required peers cannot be checked`,
        );
        continue;
      }
      issues.push(...checkRequiredPeers(manifest, installed));
    }
  }
  return [...new Set(issues)].toSorted();
};

if (import.meta.main) {
  const issues = await checkPublishedDependencies({
    emitted: process.argv.includes("--dist"),
  });
  if (issues.length) {
    process.stderr.write(`${issues.map((issue) => `- ${issue}`).join("\n")}\n`);
    process.exitCode = 1;
  } else {
    process.stdout.write(
      "Published packages declare their imports and required dependency peers.\n",
    );
  }
}
