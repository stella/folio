import { reactCompilerRules } from "@stll/oxlint-config";
import { parse } from "@vue/compiler-sfc";
import { panic } from "better-result";
import { describe, expect, test } from "bun:test";
import { readFileSync } from "node:fs";
import path from "node:path";
import ts from "typescript";
import config, {
  nonReactPackageOverride,
  reactCompilerWarningsExpireAt,
} from "../oxlint.config.ts";

const REPO_ROOT = path.resolve(import.meta.dir, "..");
const isReactModule = (specifier: string) => /^(?:react|react-dom)(?:\/|$)/u.test(specifier);
const isReactRule = (rule: string) => /^(?:react|react-hooks|react-compiler)\//u.test(rule);
const DEPENDENCY_SECTIONS = [
  "dependencies",
  "devDependencies",
  "peerDependencies",
  "optionalDependencies",
] as const;

const sourceImportsReact = (source: string, filename = "source.ts") => {
  // The SFC parser excludes templates, comments and styles from module imports.
  const scripts = [source];
  if (filename.endsWith(".vue")) {
    const { descriptor, errors } = parse(source, { filename });
    if (errors.length > 0) panic(`Cannot inspect Vue imports in ${filename}: ${errors.join("; ")}`);
    scripts.splice(
      0,
      1,
      ...[descriptor.script?.content, descriptor.scriptSetup?.content].filter(
        (script) => script !== undefined,
      ),
    );
  }
  return scripts.some((script) => {
    const file = ts.createSourceFile(filename, script, ts.ScriptTarget.Latest, true);
    let found = false;
    const visit = (node: ts.Node) => {
      if (
        (ts.isImportDeclaration(node) || ts.isExportDeclaration(node)) &&
        node.moduleSpecifier &&
        ts.isStringLiteralLike(node.moduleSpecifier) &&
        isReactModule(node.moduleSpecifier.text)
      ) {
        found = true;
      }
      if (
        ts.isImportEqualsDeclaration(node) &&
        ts.isExternalModuleReference(node.moduleReference) &&
        node.moduleReference.expression &&
        ts.isStringLiteralLike(node.moduleReference.expression) &&
        isReactModule(node.moduleReference.expression.text)
      ) {
        found = true;
      }
      if (
        ts.isImportTypeNode(node) &&
        ts.isLiteralTypeNode(node.argument) &&
        ts.isStringLiteralLike(node.argument.literal) &&
        isReactModule(node.argument.literal.text)
      ) {
        found = true;
      }
      if (
        ts.isCallExpression(node) &&
        (node.expression.kind === ts.SyntaxKind.ImportKeyword ||
          (ts.isIdentifier(node.expression) && node.expression.text === "require"))
      ) {
        const argument = node.arguments.at(0);
        if (argument && ts.isStringLiteralLike(argument) && isReactModule(argument.text))
          found = true;
      }
      ts.forEachChild(node, visit);
    };
    visit(file);
    return found;
  });
};

const packages = [...new Bun.Glob("packages/*/package.json").scanSync({ cwd: REPO_ROOT })]
  .toSorted()
  .map((manifestPath) => {
    const directory = path.dirname(manifestPath);
    const manifest = JSON.parse(readFileSync(path.join(REPO_ROOT, manifestPath), "utf8"));
    const dependsOnReact = DEPENDENCY_SECTIONS.some((section) =>
      Object.keys(manifest[section] ?? {}).some(isReactModule),
    );
    const importsReact = [
      ...new Bun.Glob("src/**/*.{ts,tsx,js,jsx,mts,cts,mjs,cjs,vue}").scanSync({
        cwd: path.join(REPO_ROOT, directory),
        onlyFiles: true,
      }),
    ].some((file) =>
      sourceImportsReact(readFileSync(path.join(REPO_ROOT, directory, file), "utf8"), file),
    );
    return { directory, usesReact: dependsOnReact || importsReact };
  });

const overrideMatches = (filename: string) =>
  nonReactPackageOverride.files.some((pattern) => new Bun.Glob(pattern).match(filename));

describe("React lint package scope", () => {
  test("React Compiler warnings expire on 2026-10-31 or #1646 merge", () => {
    // The folio lead restores errors when #1646 merges; CI caps the interim date.
    expect(reactCompilerWarningsExpireAt).toBe("2026-10-31T00:00:00.000Z");
    expect(Date.now()).toBeLessThan(Date.parse(reactCompilerWarningsExpireAt));
    expect(config.options?.denyWarnings).toBe(false);
    const compilerRules = Object.keys(reactCompilerRules);
    for (const rule of compilerRules) expect(config.rules?.[rule], rule).toBe("warn");
    for (const [rule, level] of Object.entries(config.rules ?? {})) {
      if (level === "warn") expect(compilerRules, rule).toContain(rule);
    }
  });
  test("the override covers every and only non-React package", () => {
    expect(packages.length).toBeGreaterThan(0);
    for (const { directory, usesReact } of packages) {
      expect(overrideMatches(`${directory}/src/probe.ts`), directory).toBe(!usesReact);
      expect(overrideMatches(`${directory}/src/nested/probe.tsx`), directory).toBe(!usesReact);
    }
    for (const pattern of nonReactPackageOverride.files) {
      expect(
        packages.some(({ directory }) => pattern === `${directory}/**`),
        pattern,
      ).toBe(true);
    }
  });

  test("all configured React rules are off in non-React packages", () => {
    const rules = [
      config.rules,
      ...(config.overrides ?? [])
        .filter((override) => override !== nonReactPackageOverride)
        .map((override) => override.rules),
    ];
    const reactRules = [
      ...new Set(rules.flatMap((entries) => Object.keys(entries ?? {}).filter(isReactRule))),
    ];
    expect(reactRules.length).toBeGreaterThan(0);
    expect(Object.keys(nonReactPackageOverride.rules).toSorted()).toEqual(reactRules.toSorted());
    for (const [rule, level] of Object.entries(nonReactPackageOverride.rules)) {
      expect(level, rule).toBe("off");
    }
    // Generic JSX accessibility and architecture rules retain the shared policy.
    expect(Object.keys(nonReactPackageOverride.rules).every(isReactRule)).toBe(true);
    expect(config.overrides?.at(-1)).toBe(nonReactPackageOverride);
    expect(config.rules?.["react/hooks"]).toBe("warn");
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

  test("Vue script imports count but comments and string fixtures do not", () => {
    expect(
      sourceImportsReact('<script setup lang="ts">import React from "react";</script>', "view.vue"),
    ).toBe(true);
    expect(sourceImportsReact('<template>import React from "react";</template>', "view.vue")).toBe(
      false,
    );
    expect(
      sourceImportsReact(
        '<!-- <script>import React from "react";</script> --><template />',
        "view.vue",
      ),
    ).toBe(false);
    expect(sourceImportsReact('// import React from "react";')).toBe(false);
    expect(sourceImportsReact('/* import React from "react"; */')).toBe(false);
    expect(sourceImportsReact("const fixture = 'import React from \"react\";';")).toBe(false);
    expect(sourceImportsReact('import { ref } from "vue";')).toBe(false);
  });
});
