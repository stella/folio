import { expect, test } from "bun:test";
import { Glob } from "bun";
import { readFileSync } from "node:fs";
import path from "node:path";
import ts from "typescript";
import { cloneModel } from "./modelClone";

test("model clones isolate data and retain opaque bindings across the cloned graph", () => {
  const key = Symbol("private binding");
  const handle = Object.freeze({});
  const record = { text: "value", [key]: handle };
  const source = {
    items: [record, record],
    map: new Map([[record, record]]),
    set: new Set([record]),
    bytes: new Uint8Array([1, 2]),
    date: new Date("2026-01-01"),
    expression: /value/g,
  };
  const cloned = cloneModel(source);
  const first = cloned.items.at(0);
  expect(first).toBeDefined();
  expect(first).not.toBe(record);
  expect(first?.[key]).toBe(handle);
  expect(cloned.items.at(1)).toBe(first);
  expect([...cloned.map.keys()].at(0)).toBe(first);
  expect([...cloned.map.values()].at(0)).toBe(first);
  expect([...cloned.set].at(0)).toBe(first);
  expect(cloned.bytes).toEqual(source.bytes);
  expect(cloned.bytes).not.toBe(source.bytes);
  expect(cloned.date).toEqual(source.date);
  expect(cloned.expression).toEqual(source.expression);
  if (!first) throw new TypeError("Expected cloned record");
  first.text = "edited";
  expect(record.text).toBe("value");
  expect(JSON.stringify(cloned)).not.toContain("private binding");
});

test("model clones preserve cycles and omit non-enumerable and accessor bindings", () => {
  const source: { next?: object } = {};
  source.next = source;
  const hidden = Symbol("hidden");
  const accessor = Symbol("accessor");
  Object.defineProperty(source, hidden, { value: {} });
  Object.defineProperty(source, accessor, {
    enumerable: true,
    get: () => {
      throw new TypeError("Symbol accessor must not execute");
    },
  });
  const cloned = cloneModel(source);
  expect(cloned.next).toBe(cloned);
  expect(Object.getOwnPropertySymbols(cloned)).toEqual([]);
  expect(cloneModel(null)).toBeNull();
  expect(cloneModel("value")).toBe("value");
});

const primitiveReferences = (source: string): number => {
  const file = ts.createSourceFile("fixture.ts", source, ts.ScriptTarget.Latest, true);
  let count = 0;
  const visit = (node: ts.Node): void => {
    if ((ts.isIdentifier(node) || ts.isStringLiteral(node)) && node.text === "structuredClone")
      count++;
    ts.forEachChild(node, visit);
  };
  visit(file);
  return count;
};

test("model clone ownership detects raw calls, qualified calls and aliases", () => {
  for (const source of [
    "structuredClone(value)",
    "globalThis.structuredClone(value)",
    'const copy = globalThis["structuredClone"]; copy(value)',
    "const copy = structuredClone; copy(value)",
  ]) {
    expect(primitiveReferences(source)).toBeGreaterThan(0);
  }
  expect(primitiveReferences("cloneModel(value)")).toBe(0);
});

test("every operation model clone goes through its binding owner", () => {
  const files = [...new Glob("**/*.ts").scanSync({ cwd: import.meta.dir, onlyFiles: true })].filter(
    (file) =>
      file !== "modelClone.ts" &&
      !file.includes("__tests__/") &&
      !file.endsWith(".test.ts") &&
      !file.endsWith(".typecheck.ts"),
  );
  expect(files.length).toBeGreaterThan(0);
  const violations = files.filter(
    (file) => primitiveReferences(readFileSync(path.join(import.meta.dir, file), "utf8")) > 0,
  );
  expect(violations).toEqual([]);
});
