import { expect, test } from "bun:test";
import {
  checkImports,
  checkRequiredPeers,
  type DependencyManifest,
} from "./check-package-dependencies";

const manifest = {
  name: "@example/editor",
  dependencies: { model: "^1.0.0" },
  devDependencies: { "build-types": "^1.0.0" },
  peerDependencies: { host: "^1.0.0", "optional-host": "^1.0.0" },
  peerDependenciesMeta: { "optional-host": { optional: true } },
} satisfies DependencyManifest;

test("every module-specifier syntax rejects a phantom runtime import", () => {
  const syntax = [
    'import "phantom/subpath";',
    'import { value } from "phantom";',
    'export * from "phantom";',
    'const module = import("phantom");',
    'const module = require("phantom");',
    'import module = require("phantom");',
  ];
  for (const source of syntax) {
    expect(checkImports({ manifest, source, fileName: "src/entry.ts" })).toHaveLength(1);
  }
  expect(
    checkImports({
      manifest,
      source: '<script setup lang="ts">import "phantom";</script>',
      fileName: "src/Entry.vue",
    }),
  ).toHaveLength(1);
  expect(
    checkImports({ manifest, source: '@import "phantom/styles.css";', fileName: "src/styles.css" }),
  ).toHaveLength(1);
});

test("type-only source imports allow development declarations; consumer types require runtime ownership", () => {
  for (const source of [
    'import type { Shape } from "build-types";',
    'import { type Shape } from "build-types";',
    'export type { Shape } from "build-types";',
    'export { type Shape } from "build-types";',
    'type Shape = import("build-types").Shape;',
    '/// <reference types="build-types" />',
  ]) {
    expect(checkImports({ manifest, source, fileName: "src/entry.ts" })).toEqual([]);
    expect(checkImports({ manifest, source, fileName: "dist/entry.d.ts" })).toHaveLength(1);
  }
  expect(
    checkImports({
      manifest,
      source: 'import { type Shape, value } from "build-types";',
      fileName: "src/entry.ts",
    }),
  ).toHaveLength(1);
  expect(
    checkImports({
      manifest,
      source: 'import { value } from "build-types";',
      fileName: "dist/entry.js",
    }),
  ).toHaveLength(1);
});

test("declared dependencies, peers, self imports, relative paths and builtins are accepted", () => {
  for (const specifier of [
    "model/subpath",
    "host",
    "optional-host",
    "@example/editor/subpath",
    "./local",
    "node:fs",
    "fs",
  ]) {
    expect(
      checkImports({ manifest, source: `import "${specifier}";`, fileName: "src/entry.ts" }),
    ).toEqual([]);
  }
});

test("a dependency required peer cannot be supplied by hoisting, development or optional dependencies", () => {
  const dependency = { name: "model", peerDependencies: { integration: "^1.0.0" } };
  expect(checkRequiredPeers(manifest, dependency)).toHaveLength(1);
  expect(
    checkRequiredPeers({ ...manifest, devDependencies: { integration: "^1.0.0" } }, dependency),
  ).toHaveLength(1);
  expect(
    checkRequiredPeers(
      { ...manifest, optionalDependencies: { integration: "^1.0.0" } },
      dependency,
    ),
  ).toHaveLength(1);
  expect(
    checkRequiredPeers({ ...manifest, dependencies: { integration: "^1.0.0" } }, dependency),
  ).toEqual([]);
  expect(
    checkRequiredPeers({ ...manifest, peerDependencies: { integration: "^1.0.0" } }, dependency),
  ).toEqual([]);
  expect(
    checkRequiredPeers(manifest, {
      ...dependency,
      peerDependenciesMeta: { integration: { optional: true } },
    }),
  ).toEqual([]);
});

test("scoped ambient type dependencies own corresponding declaration imports", () => {
  expect(
    checkImports({
      manifest: { name: "editor", dependencies: { "@types/example__host": "^1.0.0" } },
      source: 'import type { Host } from "@example/host";',
      fileName: "dist/entry.d.ts",
    }),
  ).toEqual([]);
});

test("source build inputs use development dependencies; emitted assets require consumer dependencies", () => {
  const stylesheet = '@import "build-types/theme.css";';
  expect(checkImports({ manifest, source: stylesheet, fileName: "src/style.css" })).toEqual([]);
  expect(checkImports({ manifest, source: stylesheet, fileName: "dist/style.css" })).toHaveLength(
    1,
  );
  const ambient = '/// <reference types="build-types" />';
  expect(checkImports({ manifest, source: ambient, fileName: "src/env.d.ts" })).toEqual([]);
  expect(checkImports({ manifest, source: ambient, fileName: "dist/env.d.ts" })).toHaveLength(1);
});
