import { afterEach, expect, test } from "bun:test";
import { mkdtempSync, mkdirSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
import {
  assertHostPeerManifest,
  assertHostPeerResolution,
  assertHostPeersExternal,
  discoverPublishedPackages,
} from "./host-peer-policy";

const temporaryDirectories: string[] = [];
afterEach(() => {
  for (const directory of temporaryDirectories.splice(0)) rmSync(directory, { recursive: true });
});
const fixture = () => {
  const root = mkdtempSync(path.join(tmpdir(), "folio-host-peers-"));
  temporaryDirectories.push(root);
  const directory = path.join(root, "packages", "new-adapter");
  mkdirSync(path.join(directory, "src"), { recursive: true });
  mkdirSync(path.join(directory, "dist"));
  writeFileSync(
    path.join(directory, "package.json"),
    JSON.stringify({ name: "@test/new-adapter" }),
  );
  return { root, directory };
};
const writeManifest = (directory: string, fields: object) => {
  writeFileSync(
    path.join(directory, "package.json"),
    JSON.stringify({ name: "@test/new-adapter", ...fields }),
  );
};

test("all discovered published manifests require host singleton peers", () => {
  const directories = discoverPublishedPackages(path.resolve(import.meta.dir, ".."));
  expect(directories.length).toBeGreaterThan(0);
  for (const directory of directories) assertHostPeerManifest(directory);
});
test("discovery includes new public packages and excludes private packages", () => {
  const { root, directory } = fixture();
  expect(discoverPublishedPackages(root)).toEqual([directory]);
  writeManifest(directory, { private: true });
  expect(discoverPublishedPackages(root)).toEqual([]);
});
test("every singleton family is rejected in either runtime dependency kind", () => {
  const { directory } = fixture();
  for (const name of [
    "prosemirror-future",
    "yjs",
    "y-prosemirror",
    "y-protocols",
    "react",
    "react-dom",
  ]) {
    for (const kind of ["dependencies", "optionalDependencies"]) {
      writeManifest(directory, { [kind]: { [name]: "^1.0.0" } });
      expect(() => assertHostPeerManifest(directory)).toThrow("must be a required peer");
    }
  }
});
test("singleton peers cannot be optional even without a source import", () => {
  const { directory } = fixture();
  writeManifest(directory, {
    peerDependencies: { yjs: "^13.6.31" },
    peerDependenciesMeta: { yjs: { optional: true } },
  });
  expect(() => assertHostPeerManifest(directory)).toThrow("must be a required peer");
});
test("new source imports require peers, including type imports and Vue scripts", () => {
  const { directory } = fixture();
  writeFileSync(
    path.join(directory, "src", "new.ts"),
    'import type { Node } from "prosemirror-model"; export type { Node };',
  );
  expect(() => assertHostPeerManifest(directory)).toThrow("undeclared peer prosemirror-model");
  writeManifest(directory, { peerDependencies: { "prosemirror-model": "^1.25.9" } });
  assertHostPeerManifest(directory);
  writeFileSync(
    path.join(directory, "src", "new.vue"),
    '<script setup>import * as Y from "yjs";</script>',
  );
  expect(() => assertHostPeerManifest(directory)).toThrow("undeclared peer yjs");
});
test("test imports do not create production peer requirements", () => {
  const { directory } = fixture();
  writeFileSync(path.join(directory, "src", "editor.test.ts"), 'import "prosemirror-model";');
  assertHostPeerManifest(directory);
});
test("runtime subpaths stay external while type imports are erased", () => {
  const { directory } = fixture();
  writeManifest(directory, {
    peerDependencies: { "y-protocols": "^1.0.6", "prosemirror-model": "^1.25.9" },
  });
  writeFileSync(
    path.join(directory, "src", "index.ts"),
    'import type { Node } from "prosemirror-model"; export type { Node }; export { Awareness } from "y-protocols/awareness";',
  );
  writeFileSync(
    path.join(directory, "dist", "index.js"),
    'export { Awareness } from "y-protocols/awareness";',
  );
  assertHostPeersExternal(directory);
  writeFileSync(
    path.join(directory, "dist", "index.js"),
    "class Awareness {} export { Awareness };",
  );
  expect(() => assertHostPeersExternal(directory)).toThrow("missing from built externals");
});
test("bundled internals fail even alongside a retained bare import", () => {
  const { directory } = fixture();
  writeManifest(directory, { peerDependencies: { "prosemirror-state": "^1.4.4" } });
  const imported = 'export { PluginKey } from "prosemirror-state";';
  writeFileSync(path.join(directory, "src", "index.ts"), imported);
  writeFileSync(
    path.join(directory, "dist", "index.js"),
    `${imported}\nthrow new RangeError("Adding different instances of a keyed plugin");`,
  );
  expect(() => assertHostPeersExternal(directory)).toThrow("bundled singleton code");
  writeFileSync(path.join(directory, "dist", "index.js"), imported);
  writeFileSync(
    path.join(directory, "dist", "index.js.map"),
    JSON.stringify({ sources: ["../../node_modules/prosemirror-state/dist/index.js"] }),
  );
  expect(() => assertHostPeersExternal(directory)).toThrow("bundled singleton source");
});

test("installed peers resolve to the host copy and reject nested private copies", () => {
  for (const nested of [false, true]) {
    const { root } = fixture();
    writeFileSync(path.join(root, "package.json"), "{}");
    const adapter = path.join(root, "node_modules", "@stll", "folio-new");
    mkdirSync(adapter, { recursive: true });
    writeFileSync(
      path.join(adapter, "package.json"),
      JSON.stringify({ name: "@stll/folio-new", peerDependencies: { "y-protocols": "^1.0.6" } }),
    );
    const directories = [path.join(root, "node_modules", "y-protocols")];
    if (nested) directories.push(path.join(adapter, "node_modules", "y-protocols"));
    for (const directory of directories) {
      mkdirSync(directory, { recursive: true });
      writeFileSync(
        path.join(directory, "package.json"),
        JSON.stringify({ name: "y-protocols", exports: { "./awareness": "./awareness.cjs" } }),
      );
      writeFileSync(path.join(directory, "awareness.cjs"), "module.exports = {};");
    }
    if (nested) expect(() => assertHostPeerResolution(root)).toThrow("host uses");
    else assertHostPeerResolution(root);
  }
});
