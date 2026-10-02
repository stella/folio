import { expect, test } from "bun:test";
import { existsSync, readFileSync, readdirSync } from "node:fs";
import { mkdir, mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import ts from "typescript";

import { getFolioMessages } from "../packages/core/src/i18n/messages";
import { FOLIO_LOCALES } from "../packages/core/src/i18n/messages/locales";

const root = path.resolve(import.meta.dir, "..");
const packages = ["core", "react", "vue"] as const;
const localeFilename = (packageName: string, locale: string) =>
  `${locale}${packageName === "core" ? ".gen" : ""}.ts`;
const messagesDirectory = (packageName: string) =>
  path.join(root, "packages", packageName, "src/i18n/messages");

const runtimeSpecifiers = (file: string) => {
  const source = ts.createSourceFile(
    file,
    readFileSync(file, "utf8"),
    ts.ScriptTarget.Latest,
    true,
  );
  const specifiers: string[] = [];
  const visit = (node: ts.Node) => {
    if (ts.isImportDeclaration(node) && ts.isStringLiteral(node.moduleSpecifier)) {
      const clause = node.importClause;
      const bindings = clause?.namedBindings;
      const onlyTypes =
        clause?.isTypeOnly ||
        (!clause?.name &&
          bindings &&
          ts.isNamedImports(bindings) &&
          bindings.elements.length > 0 &&
          bindings.elements.every((binding) => binding.isTypeOnly));
      if (!onlyTypes) specifiers.push(node.moduleSpecifier.text);
    }
    if (
      ts.isExportDeclaration(node) &&
      node.moduleSpecifier &&
      ts.isStringLiteral(node.moduleSpecifier) &&
      !node.isTypeOnly
    ) {
      const onlyTypes =
        node.exportClause &&
        ts.isNamedExports(node.exportClause) &&
        node.exportClause.elements.length > 0 &&
        node.exportClause.elements.every((binding) => binding.isTypeOnly);
      if (!onlyTypes) specifiers.push(node.moduleSpecifier.text);
    }
    if (
      ts.isCallExpression(node) &&
      (node.expression.kind === ts.SyntaxKind.ImportKeyword ||
        (ts.isIdentifier(node.expression) && node.expression.text === "require"))
    ) {
      const argument = node.arguments.at(0);
      // A computed dependency defeats this source graph's isolation proof.
      expect(argument && ts.isStringLiteral(argument)).toBe(true);
      if (argument && ts.isStringLiteral(argument)) specifiers.push(argument.text);
    }
    ts.forEachChild(node, visit);
  };
  visit(source);
  return specifiers;
};

const runtimeGraph = (entry: string) => {
  const visited = new Set<string>();
  const pending = [entry];
  while (pending.length > 0) {
    const file = pending.pop();
    if (!file || visited.has(file)) continue;
    visited.add(file);
    for (const specifier of runtimeSpecifiers(file)) {
      let base: string;
      if (specifier.startsWith(".")) base = path.resolve(path.dirname(file), specifier);
      else if (specifier.startsWith("@stll/folio-core/")) {
        base = path.join(root, "packages/core/src", specifier.slice("@stll/folio-core/".length));
        if (path.dirname(base) === messagesDirectory("core") && path.basename(base) !== "locales")
          base += ".gen";
      } else {
        // These data-only entry points have no runtime dependency on external packages.
        expect(specifier).toBe("");
        continue;
      }
      const resolved = [base, `${base}.ts`, base.replace(/\.js$/u, ".ts")].find(existsSync);
      expect(resolved).toBeDefined();
      if (resolved) pending.push(resolved);
    }
  }
  return visited;
};

test("generated entry sets exactly cover the shipped catalogs and lightweight locale list", () => {
  const expected = [...FOLIO_LOCALES, "locales"].sort();
  expect(new Set(FOLIO_LOCALES).size).toBe(FOLIO_LOCALES.length);
  const catalogs = readdirSync(messagesDirectory("core"))
    .filter((file) => file.endsWith(".json"))
    .map((file) => file.slice(0, -5))
    .sort();
  expect(catalogs).toEqual([...FOLIO_LOCALES].sort());
  for (const packageName of packages) {
    const entries = readdirSync(messagesDirectory(packageName))
      .filter(
        (file) => file.endsWith(".ts") && file !== "catalogs.gen.ts" && file !== "messages.gen.ts",
      )
      .map((file) => file.replace(/(?:\.gen)?\.ts$/u, ""))
      .sort();
    expect(entries).toEqual(expected);
  }
});

for (const packageName of packages) {
  for (const locale of FOLIO_LOCALES) {
    test(`${packageName}/${locale} exports the existing catalog without other locales`, async () => {
      const entry = path.join(messagesDirectory(packageName), localeFilename(packageName, locale));
      const module = await import(entry);
      expect(module.default).toBe(module.messages);
      expect(module.messages).toEqual(getFolioMessages(locale));
      expect(module.messages).toEqual(
        JSON.parse(readFileSync(path.join(messagesDirectory("core"), `${locale}.json`), "utf8")),
      );
      for (const dependency of runtimeGraph(entry)) {
        expect(dependency).not.toBe(path.join(root, "packages/core/src/i18n/messages.ts"));
        expect(path.basename(dependency)).not.toBe("catalogs.gen.ts");
        if (path.dirname(dependency) === messagesDirectory("core")) {
          expect([`${locale}.gen.ts`, "locales.ts"]).toContain(path.basename(dependency));
        }
      }
    });
  }
  test(`${packageName}/locales is a catalog-free locale discovery entry`, async () => {
    const entry = path.join(messagesDirectory(packageName), "locales.ts");
    expect((await import(entry)).FOLIO_LOCALES).toEqual(FOLIO_LOCALES);
    for (const dependency of runtimeGraph(entry)) {
      expect(path.basename(dependency)).toBe("locales.ts");
    }
  });
}

test("public export patterns resolve each generated entry", async () => {
  for (const packageName of packages) {
    const manifest = await Bun.file(
      path.join(root, "packages", packageName, "package.json"),
    ).json();
    const pattern = packageName === "core" ? "./i18n/messages/*" : "./messages/*";
    if (packageName === "vue") {
      expect(manifest.exports[pattern]).toEqual({
        types: "./dist/i18n/messages/*.d.ts",
        import: "./dist/messages/*.js",
        require: "./dist/messages/*.cjs",
      });
      continue;
    }
    expect(manifest.exports[pattern]).toBe(
      packageName === "core" ? "./src/i18n/messages/*.gen.ts" : "./src/i18n/messages/*.ts",
    );
    if (packageName === "core")
      expect(manifest.exports["./i18n/messages/locales"]).toBe("./src/i18n/messages/locales.ts");
    for (const locale of [...FOLIO_LOCALES, "locales"]) {
      expect(
        existsSync(
          path.join(
            root,
            "packages",
            packageName,
            locale === "locales" && packageName === "core"
              ? manifest.exports["./i18n/messages/locales"]
              : manifest.exports[pattern].replace("*", locale),
          ),
        ),
      ).toBe(true);
    }
  }
});

test("locale exports preserve the generated message-key type used by React", () => {
  const manifest = JSON.parse(readFileSync(path.join(root, "packages/core/package.json"), "utf8"));
  expect(manifest.exports["./i18n/messages/*.gen"]).toBe("./src/i18n/messages/*.gen.ts");
  // Node resolves exports without Bun's workspace tsconfig path aliases.
  const result = Bun.spawnSync(
    [
      "node",
      "--input-type=module",
      "-e",
      'import { createRequire } from "node:module"; process.stdout.write(createRequire(`${process.cwd()}/package.json`).resolve("@stll/folio-core/i18n/messages/messages.gen"));',
    ],
    { cwd: path.join(root, "packages/react"), stdout: "pipe", stderr: "pipe" },
  );
  expect(result.exitCode).toBe(0);
  expect(result.stdout.toString()).toBe(path.join(messagesDirectory("core"), "messages.gen.ts"));
});

test("build configurations include locale-derived entries and the discovery module", () => {
  const coreConfig = readFileSync(path.join(root, "packages/core/tsdown.config.ts"), "utf8");
  expect(coreConfig).toContain('"src/**/*.ts"');
  expect(coreConfig).not.toMatch(/![^"\n]*i18n\/messages/u);
  for (const packageName of ["react", "vue"]) {
    const configFile = path.join(
      root,
      "packages",
      packageName,
      packageName === "react" ? "tsdown.config.ts" : "vite.config.ts",
    );
    const imports = runtimeSpecifiers(configFile);
    expect(imports).toContain("@stll/folio-core/i18n/messages/locales");
    expect(imports.some((specifier) => specifier.startsWith("../core/"))).toBe(false);
    const source = ts.createSourceFile(
      configFile,
      readFileSync(configFile, "utf8"),
      ts.ScriptTarget.Latest,
      true,
    );
    const prefix = packageName === "react" ? "i18n/messages" : "messages";
    const entrySpreads: string[] = [];
    const visit = (node: ts.Node) => {
      if (ts.isSpreadAssignment(node)) entrySpreads.push(node.expression.getText(source));
      ts.forEachChild(node, visit);
    };
    visit(source);
    expect(
      entrySpreads.some(
        (expression) =>
          expression.includes("Object.fromEntries") &&
          expression.includes("FOLIO_LOCALES") &&
          expression.includes(`${prefix}/`) &&
          expression.includes("src/i18n/messages/"),
      ),
    ).toBe(true);
    // The discovery entry may be explicitly declared or included in the mapped names.
    const runtimeText = source.statements.map((statement) => statement.getText(source)).join("\n");
    expect(runtimeText).toContain("locales");
  }
});

test("publish preparation preserves nested locale wildcard destinations", async () => {
  const packageRoot = await mkdtemp(path.join(tmpdir(), "folio-locale-publish-"));
  try {
    await mkdir(path.join(packageRoot, "dist"));
    await writeFile(path.join(packageRoot, "dist/index.js"), "export {};\n");
    await writeFile(path.join(packageRoot, "dist/index.d.ts"), "export {};\n");
    for (const pattern of ["./messages/*", "./i18n/messages/*"]) {
      const suffix = pattern === "./i18n/messages/*" ? ".gen" : "";
      await writeFile(
        path.join(packageRoot, "package.json"),
        JSON.stringify({
          name: "@stll/locale-fixture",
          version: "1.0.0",
          exports: { ".": "./src/index.ts", [pattern]: `./src/i18n/messages/*${suffix}.ts` },
        }),
      );
      const result = Bun.spawnSync(
        [process.execPath, path.join(import.meta.dir, "prepare-publish.ts"), packageRoot],
        { stderr: "pipe", stdout: "pipe" },
      );
      expect(result.exitCode).toBe(0);
      const manifest = await Bun.file(path.join(packageRoot, "package.json")).json();
      expect(manifest.exports[pattern]).toEqual({
        types: `./dist/i18n/messages/*${suffix}.d.ts`,
        import: `./dist/i18n/messages/*${suffix}.js`,
      });
    }
  } finally {
    await rm(packageRoot, { recursive: true, force: true });
  }
});
