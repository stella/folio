import { expect, test } from "bun:test";
import { existsSync, readdirSync, readFileSync } from "node:fs";
import { mkdir, mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import type { StringLiteral } from "oxc-parser";

import { getFolioMessages } from "../packages/core/src/i18n/messages";
import { FOLIO_LOCALES } from "../packages/core/src/i18n/messages/locales";
import { type Node, nodeText, parseSourceModule, walkAst } from "./lib/oxc-ast";

const root = path.resolve(import.meta.dir, "..");
const packages = ["core", "react", "vue"] as const;
const localeFilename = (packageName: string, locale: string) =>
  `${locale}${packageName === "core" ? ".gen" : ""}.ts`;
const messagesDirectory = (packageName: string) =>
  path.join(root, "packages", packageName, "src/i18n/messages");

const isStringLiteral = (node: Node): node is StringLiteral =>
  node.type === "Literal" && typeof node.value === "string";

const runtimeSpecifiers = (file: string) => {
  const { program } = parseSourceModule(file, readFileSync(file, "utf8"));
  const specifiers: string[] = [];
  walkAst(program, (node) => {
    if (node.type === "ImportDeclaration") {
      const onlyTypes =
        node.importKind === "type" ||
        (node.specifiers.length > 0 &&
          node.specifiers.every(
            (binding) =>
              binding.type === "ImportSpecifier" &&
              binding.importKind === "type",
          ));
      if (!onlyTypes) specifiers.push(node.source.value);
    }
    if (
      (node.type === "ExportNamedDeclaration" ||
        node.type === "ExportAllDeclaration") &&
      node.source &&
      node.exportKind !== "type"
    ) {
      const onlyTypes =
        node.type === "ExportNamedDeclaration" &&
        node.specifiers.length > 0 &&
        node.specifiers.every((binding) => binding.exportKind === "type");
      if (!onlyTypes) specifiers.push(node.source.value);
    }
    if (
      node.type === "ImportExpression" ||
      (node.type === "CallExpression" &&
        node.callee.type === "Identifier" &&
        node.callee.name === "require")
    ) {
      const argument =
        node.type === "ImportExpression" ? node.source : node.arguments.at(0);
      // A computed dependency defeats this source graph's isolation proof.
      expect(argument && isStringLiteral(argument)).toBe(true);
      if (argument && isStringLiteral(argument))
        specifiers.push(argument.value);
    }
  });
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
      if (specifier.startsWith("."))
        base = path.resolve(path.dirname(file), specifier);
      else if (specifier.startsWith("@stll/folio-core/")) {
        base = path.join(
          root,
          "packages/core/src",
          specifier.slice("@stll/folio-core/".length),
        );
        if (
          path.dirname(base) === messagesDirectory("core") &&
          path.basename(base) !== "locales"
        )
          base += ".gen";
      } else {
        // These data-only entry points have no runtime dependency on external packages.
        expect(specifier).toBe("");
        continue;
      }
      const resolved = [base, `${base}.ts`, base.replace(/\.js$/u, ".ts")].find(
        existsSync,
      );
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
        (file) =>
          file.endsWith(".ts") &&
          file !== "catalogs.gen.ts" &&
          file !== "messages.gen.ts",
      )
      .map((file) => file.replace(/(?:\.gen)?\.ts$/u, ""))
      .sort();
    expect(entries).toEqual(expected);
  }
});

for (const packageName of packages) {
  for (const locale of FOLIO_LOCALES) {
    test(`${packageName}/${locale} exports the existing catalog without other locales`, async () => {
      const entry = path.join(
        messagesDirectory(packageName),
        localeFilename(packageName, locale),
      );
      const module = await import(entry);
      expect(module.default).toBe(module.messages);
      expect(module.messages).toEqual(getFolioMessages(locale));
      expect(module.messages).toEqual(
        JSON.parse(
          readFileSync(
            path.join(messagesDirectory("core"), `${locale}.json`),
            "utf8",
          ),
        ),
      );
      for (const dependency of runtimeGraph(entry)) {
        expect(dependency).not.toBe(
          path.join(root, "packages/core/src/i18n/messages.ts"),
        );
        expect(path.basename(dependency)).not.toBe("catalogs.gen.ts");
        if (path.dirname(dependency) === messagesDirectory("core")) {
          expect([`${locale}.gen.ts`, "locales.ts"]).toContain(
            path.basename(dependency),
          );
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
    const pattern =
      packageName === "core" ? "./i18n/messages/*" : "./messages/*";
    if (packageName === "vue") {
      expect(manifest.exports[pattern]).toEqual({
        types: "./dist/i18n/messages/*.d.ts",
        import: "./dist/messages/*.js",
        require: "./dist/messages/*.cjs",
      });
      continue;
    }
    expect(manifest.exports[pattern]).toBe(
      packageName === "core"
        ? "./src/i18n/messages/*.gen.ts"
        : "./src/i18n/messages/*.ts",
    );
    if (packageName === "core")
      expect(manifest.exports["./i18n/messages/locales"]).toBe(
        "./src/i18n/messages/locales.ts",
      );
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
  const manifest = JSON.parse(
    readFileSync(path.join(root, "packages/core/package.json"), "utf8"),
  );
  expect(manifest.exports["./i18n/messages/*.gen"]).toBe(
    "./src/i18n/messages/*.gen.ts",
  );
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
  expect(result.stdout.toString()).toBe(
    path.join(messagesDirectory("core"), "messages.gen.ts"),
  );
});

test("build configurations include locale-derived entries and the discovery module", () => {
  const coreConfig = readFileSync(
    path.join(root, "packages/core/tsdown.config.ts"),
    "utf8",
  );
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
    expect(imports.some((specifier) => specifier.startsWith("../core/"))).toBe(
      false,
    );
    const { program, text } = parseSourceModule(
      configFile,
      readFileSync(configFile, "utf8"),
    );
    const prefix = packageName === "react" ? "i18n/messages" : "messages";
    const entrySpreads: string[] = [];
    walkAst(program, (node, ancestors) => {
      if (
        node.type === "SpreadElement" &&
        ancestors.at(-1)?.type === "ObjectExpression"
      )
        entrySpreads.push(nodeText(text, node.argument));
    });
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
    const runtimeText = program.body
      .map((statement) => nodeText(text, statement))
      .join("\n");
    expect(runtimeText).toContain("locales");
  }
});

test("publish preparation preserves nested locale wildcard destinations", async () => {
  const packageRoot = await mkdtemp(
    path.join(tmpdir(), "folio-locale-publish-"),
  );
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
          exports: {
            ".": "./src/index.ts",
            [pattern]: `./src/i18n/messages/*${suffix}.ts`,
          },
        }),
      );
      const result = Bun.spawnSync(
        [
          process.execPath,
          path.join(import.meta.dir, "prepare-publish.ts"),
          packageRoot,
        ],
        { stderr: "pipe", stdout: "pipe" },
      );
      expect(result.exitCode).toBe(0);
      const manifest = await Bun.file(
        path.join(packageRoot, "package.json"),
      ).json();
      expect(manifest.exports[pattern]).toEqual({
        types: `./dist/i18n/messages/*${suffix}.d.ts`,
        import: `./dist/i18n/messages/*${suffix}.js`,
        default: `./dist/i18n/messages/*${suffix}.js`,
      });
    }
  } finally {
    await rm(packageRoot, { recursive: true, force: true });
  }
});
