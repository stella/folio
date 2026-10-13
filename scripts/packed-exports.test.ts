import { expect, test } from "bun:test";
import { publicEntries } from "./lib/packed-exports";

test("every packed wildcard target is exercised, including CJS and declarations", () => {
  const exports = {
    ".": { types: "./dist/index.d.ts", import: "./dist/index.js", require: "./dist/index.cjs" },
    "./*": { types: "./dist/*.d.ts", import: "./dist/*.js", require: "./dist/*.cjs" },
    "./internal/*": null,
    "./private": null,
    "./editor.css": "./dist/editor.css",
  };
  const files = [
    "dist/index.js",
    "dist/index.cjs",
    "dist/index.d.ts",
    "dist/nested/public.js",
    "dist/nested/public.cjs",
    "dist/nested/public.d.ts",
    "dist/private.js",
    "dist/internal/hidden.js",
    "dist/editor.css",
  ];
  const entries = publicEntries({ name: "@example/editor", exports, files });
  expect(entries.map(({ specifier }) => specifier)).toEqual([
    "@example/editor",
    "@example/editor/editor.css",
    "@example/editor/index",
    "@example/editor/nested/public",
  ]);
  expect(entries.filter(({ cjs }) => cjs)).toHaveLength(3);
  expect(entries.filter(({ types }) => types)).toHaveLength(3);
  expect(entries.filter(({ asset }) => asset)).toHaveLength(1);
});

test("exact exports and more specific patterns override broad wildcard exports", () => {
  const entries = publicEntries({
    name: "example",
    exports: {
      "./*": { import: "./dist/*.js" },
      "./messages/*": { import: "./dist/i18n/*.js" },
      "./messages/en": null,
    },
    files: ["dist/messages/en.js", "dist/i18n/en.js", "dist/i18n/cs.js"],
  });
  expect(entries.map(({ specifier }) => specifier)).toEqual([
    "example/i18n/cs",
    "example/i18n/en",
    "example/messages/cs",
  ]);
});

test("plain JavaScript exports remain public even without a types condition", () => {
  const entries = publicEntries({
    name: "example",
    exports: {
      ".": "./index.js",
      "./commonjs": "./commonjs.cjs",
    },
    files: ["index.js", "commonjs.cjs"],
  });
  expect(entries.map(({ specifier }) => specifier)).toEqual(["example", "example/commonjs"]);
  expect(entries.every(({ asset }) => !asset)).toBe(true);
  expect(entries.find(({ specifier }) => specifier.endsWith("/commonjs"))?.cjs).toBe(true);
});
