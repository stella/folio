import { parse } from "@vue/compiler-sfc";
import { panic } from "better-result";
import { describe, expect, test } from "bun:test";
import {
  mkdirSync,
  mkdtempSync,
  readFileSync,
  rmSync,
  writeFileSync,
} from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";

import { reactCompilerRules } from "@stll/oxlint-config";

import config, { nonReactPackageOverride } from "../oxlint.config.ts";
import { type Node, parseSourceModule, walkAst } from "./lib/oxc-ast";

const REPO_ROOT = path.resolve(import.meta.dir, "..");
const isReactModule = (specifier: string) =>
  /^(?:react|react-dom)(?:\/|$)/u.test(specifier);
const isReactRule = (rule: string) =>
  /^(?:react|react-hooks|react-compiler)\//u.test(rule);
/** A string literal or a template literal without substitutions. */
const stringLiteralText = (node: Node | undefined): string | undefined => {
  if (node?.type === "Literal") {
    const value: unknown = node.value;
    return typeof value === "string" ? value : undefined;
  }
  if (node?.type === "TemplateLiteral" && node.expressions.length === 0) {
    return node.quasis[0]?.value.cooked ?? undefined;
  }
  return undefined;
};
/** The key of a plain `key: value` property with an identifier or string key. */
const propertyAssignmentName = (node: Node): string | undefined => {
  if (
    node.type !== "Property" ||
    node.kind !== "init" ||
    node.method ||
    node.shorthand ||
    node.computed
  ) {
    return undefined;
  }
  return node.key.type === "Identifier"
    ? node.key.name
    : stringLiteralText(node.key);
};
const sourceImportsReact = (source: string, filename = "source.ts") => {
  // The SFC parser excludes templates, comments and styles from module imports.
  const scripts = [source];
  if (filename.endsWith(".vue")) {
    const { descriptor, errors } = parse(source, { filename });
    if (errors.length > 0)
      panic(`Cannot inspect Vue imports in ${filename}: ${errors.join("; ")}`);
    scripts.splice(
      0,
      1,
      ...[descriptor.script?.content, descriptor.scriptSetup?.content].filter(
        (script) => script !== undefined,
      ),
    );
  }
  // Vue scripts are TypeScript-shaped; parse them as TypeScript.
  const scriptName = filename.endsWith(".vue") ? `${filename}.ts` : filename;
  return scripts.some((script) => {
    const { program } = parseSourceModule(scriptName, script);
    let found = false;
    walkAst(program, (node) => {
      if (
        (node.type === "ImportDeclaration" ||
          node.type === "ExportNamedDeclaration" ||
          node.type === "ExportAllDeclaration") &&
        node.source &&
        isReactModule(node.source.value)
      ) {
        found = true;
      }
      if (
        node.type === "TSImportEqualsDeclaration" &&
        node.moduleReference.type === "TSExternalModuleReference" &&
        isReactModule(node.moduleReference.expression.value)
      ) {
        found = true;
      }
      if (node.type === "TSImportType" && isReactModule(node.source.value)) {
        found = true;
      }
      let argument: Node | undefined;
      if (node.type === "ImportExpression") argument = node.source;
      else if (
        node.type === "CallExpression" &&
        node.callee.type === "Identifier" &&
        node.callee.name === "require"
      )
        argument = node.arguments.at(0);
      const specifier = stringLiteralText(argument);
      if (specifier !== undefined && isReactModule(specifier)) found = true;
    });
    return found;
  });
};

const packageImportsReact = (directory: string) =>
  [
    ...new Bun.Glob("src/**/*.{ts,tsx,js,jsx,mts,cts,mjs,cjs,vue}").scanSync({
      cwd: directory,
      onlyFiles: true,
    }),
  ].some((file) =>
    sourceImportsReact(readFileSync(path.join(directory, file), "utf8"), file),
  );

const packages = [
  ...new Bun.Glob("packages/*/package.json").scanSync({ cwd: REPO_ROOT }),
]
  .toSorted()
  .map((manifestPath) => {
    const directory = path.dirname(manifestPath);
    const importsReact = packageImportsReact(path.join(REPO_ROOT, directory));
    // A required peer of a framework-neutral dependency does not make Vue
    // source React-owned. Apply hook rules according to the source import census.
    return { directory, usesReact: importsReact };
  });

const overrideMatches = (filename: string) =>
  nonReactPackageOverride.files.some((pattern) =>
    new Bun.Glob(pattern).match(filename),
  );
const reactPresetOverride = config.overrides?.find((override) =>
  Object.hasOwn(override.rules ?? {}, "react/hooks"),
);
const reactPresetMatches = (filename: string) =>
  reactPresetOverride?.files.some((pattern) =>
    new Bun.Glob(pattern).match(filename),
  ) ?? false;

describe("React lint package scope", () => {
  test("the shared config owns React Compiler policy without local severity overrides", () => {
    expect(config.options?.denyWarnings).toBe(true);
    const compilerRules = Object.keys(reactCompilerRules);
    expect(compilerRules.length).toBeGreaterThan(0);
    expect(reactPresetOverride).toBeDefined();
    expect(reactPresetOverride?.plugins).toContain("react");
    for (const [rule, level] of Object.entries(reactCompilerRules)) {
      if (level !== "off") expect(level, rule).toBe("error");
      expect(reactPresetOverride?.rules?.[rule], rule).toBe(level);
    }
    const source = readFileSync(
      path.join(REPO_ROOT, "oxlint.config.ts"),
      "utf8",
    );
    expect(source).not.toMatch(
      /WarningsExpireAt|WarningPolicyAt|Date\.(?:now|parse)\(/u,
    );
    const { program } = parseSourceModule("oxlint.config.ts", source);
    const declaration = program.body.find(
      (statement) => statement.type === "ExportDefaultDeclaration",
    );
    if (declaration?.declaration.type !== "CallExpression") {
      panic("Lint config must export the shared library config call");
    }
    const options = declaration.declaration.arguments.at(0);
    if (options?.type !== "ObjectExpression") {
      panic("Lint config options must be inspectable");
    }
    const rulesProperty = options.properties.find(
      (property) => propertyAssignmentName(property) === "rules",
    );
    if (
      rulesProperty?.type !== "Property" ||
      rulesProperty.value.type !== "ObjectExpression"
    ) {
      panic("Local lint rules must be an explicit object");
    }
    for (const property of rulesProperty.value.properties) {
      const name = propertyAssignmentName(property);
      if (name === undefined) {
        panic(
          "Local rule spreads and computed keys can override shared Compiler policy",
        );
      }
      expect(compilerRules, name).not.toContain(name);
    }
    expect(config.rules).not.toHaveProperty("react/hooks");
    for (const override of config.overrides ?? []) {
      if (
        override === reactPresetOverride ||
        override === nonReactPackageOverride
      )
        continue;
      for (const rule of Object.keys(override.rules ?? {})) {
        expect(compilerRules, rule).not.toContain(rule);
      }
    }
  });

  test("lint invocations leave warning policy to the config", () => {
    const paths = new Set([
      "package.json",
      "lefthook.yml",
      ...new Bun.Glob("packages/**/package.json").scanSync({ cwd: REPO_ROOT }),
      ...new Bun.Glob(".github/**/*.{yml,yaml,sh,js,ts,mjs,cjs}").scanSync({
        cwd: REPO_ROOT,
      }),
      ...new Bun.Glob("scripts/**/*.{sh,js,ts,mjs,cjs}").scanSync({
        cwd: REPO_ROOT,
      }),
      ...new Bun.Glob(".husky/**").scanSync({
        cwd: REPO_ROOT,
        onlyFiles: true,
      }),
    ]);
    for (const file of paths) {
      // This guard contains flag fixtures, not lint invocations.
      if (file === "scripts/react-lint-scope.test.ts") continue;
      const source = readFileSync(path.join(REPO_ROOT, file), "utf8");
      if (/\b(?:oxlint|lint(?::[\w-]+)?)\b/u.test(source)) {
        expect(source, file).not.toMatch(/--(?:deny-warnings|max-warnings)\b/u);
      }
    }
  });

  test("React and non-React overrides match the package import census", () => {
    expect(packages.length).toBeGreaterThan(0);
    for (const { directory, usesReact } of packages) {
      expect(reactPresetMatches(`${directory}/src/probe.ts`), directory).toBe(
        usesReact,
      );
      expect(
        reactPresetMatches(`${directory}/src/nested/probe.tsx`),
        directory,
      ).toBe(usesReact);
      expect(overrideMatches(`${directory}/src/probe.ts`), directory).toBe(
        !usesReact,
      );
      expect(
        overrideMatches(`${directory}/src/nested/probe.tsx`),
        directory,
      ).toBe(!usesReact);
    }
    for (const pattern of nonReactPackageOverride.files) {
      expect(
        packages.some(({ directory }) => pattern === `${directory}/**`),
        pattern,
      ).toBe(true);
    }
  });

  test("React Compiler policy is confined to React-owned source", () => {
    expect(reactPresetOverride).toBeDefined();
    expect(
      Object.keys(reactPresetOverride?.rules ?? {}).filter(isReactRule).length,
    ).toBeGreaterThan(0);
    for (const { directory, usesReact } of packages) {
      expect(reactPresetMatches(`${directory}/src/probe.ts`), directory).toBe(
        usesReact,
      );
    }
    expect(config.overrides?.at(-1)).toBe(nonReactPackageOverride);
    expect(reactPresetOverride?.rules?.["react/hooks"]).toBe(
      reactCompilerRules["react/hooks"],
    );
  });

  test.each([
    'import React from "react";',
    'import type { ReactNode } from "react";',
    'import "react";',
    'export { useState } from "react";',
    'const react = import("react/jsx-runtime");',
    'const react = require("react");',
    'import React = require("react");',
    "const react = import(`react`);",
    'type Node = import("react").ReactNode;',
    'import { createRoot } from "react-dom/client";',
  ])("detects real React imports: %s", (source) => {
    expect(sourceImportsReact(source)).toBe(true);
  });

  test("required React peers do not change a framework-neutral source package's lint ownership", () => {
    const root = mkdtempSync(path.join(tmpdir(), "folio-react-source-census-"));
    try {
      mkdirSync(path.join(root, "src"));
      writeFileSync(
        path.join(root, "package.json"),
        JSON.stringify({
          name: "framework-neutral-adapter",
          dependencies: { react: "^19.2.0" },
        }),
      );
      const entry = path.join(root, "src", "entry.ts");
      writeFileSync(entry, 'import { createTranslator } from "use-intl/core";');
      expect(packageImportsReact(root)).toBe(false);
      writeFileSync(entry, 'import { useState } from "react";');
      expect(packageImportsReact(root)).toBe(true);
    } finally {
      rmSync(root, { recursive: true, force: true });
    }
  });

  test("Vue script imports count but comments and string fixtures do not", () => {
    expect(
      sourceImportsReact(
        '<script setup lang="ts">import React from "react";</script>',
        "view.vue",
      ),
    ).toBe(true);
    expect(
      sourceImportsReact(
        '<template>import React from "react";</template>',
        "view.vue",
      ),
    ).toBe(false);
    expect(
      sourceImportsReact(
        '<!-- <script>import React from "react";</script> --><template />',
        "view.vue",
      ),
    ).toBe(false);
    expect(sourceImportsReact('// import React from "react";')).toBe(false);
    expect(sourceImportsReact('/* import React from "react"; */')).toBe(false);
    expect(
      sourceImportsReact("const fixture = 'import React from \"react\";';"),
    ).toBe(false);
    expect(sourceImportsReact('import { ref } from "vue";')).toBe(false);
  });
});
