import { describe, expect, test } from "bun:test";
import { writeFile } from "node:fs/promises";
import path from "node:path";

import { makeTempDir } from "./__tests__/fixtures";
import { coerceFlag, flagsForSchema, readInputSource, type GeneratedFlag } from "./flags";
import { findFileTool } from "./registry";

const flag = (valueType: GeneratedFlag["valueType"]): GeneratedFlag => ({
  property: "value",
  flag: "value",
  valueType,
  description: "",
});

describe("flagsForSchema", () => {
  test("derives kebab-case flags typed by each property's schema", () => {
    const tool = findFileTool("find_text");
    if (!tool) throw new Error("find_text is not registered");

    const flags = flagsForSchema(tool.argsSchema);

    expect(flags.map(({ flag: name, valueType }) => [name, valueType])).toEqual([
      ["query", "string"],
      ["match-case", "boolean"],
      ["whole-word", "boolean"],
      ["scope", "json"],
    ]);
  });
});

describe("coerceFlag", () => {
  test("parses numbers, integers and JSON, and refuses mismatched values", () => {
    expect(coerceFlag({ flag: flag("integer"), raw: "12" }).unwrap()).toBe(12);
    expect(coerceFlag({ flag: flag("number"), raw: "1.5" }).unwrap()).toBe(1.5);
    expect(coerceFlag({ flag: flag("json"), raw: '{"a":1}' }).unwrap()).toEqual({ a: 1 });
    expect(coerceFlag({ flag: flag("boolean"), raw: true }).unwrap()).toBe(true);

    for (const [valueType, raw] of [
      ["integer", "1.5"],
      ["integer", ""],
      ["number", "ten"],
      ["json", "{"],
      ["boolean", "yes"],
      ["string", true],
    ] as const) {
      const result = coerceFlag({ flag: flag(valueType), raw });
      expect(result.isErr() && result.error.code).toBe("usage_error");
    }
  });
});

describe("readInputSource", () => {
  test("reads inline JSON, @file, and stdin", async () => {
    const { dir, cleanup } = await makeTempDir();
    try {
      const file = path.join(dir, "args.json");
      await writeFile(file, '{"query":"from file"}');
      const readStdin = () => Promise.resolve('[{"type":"deleteBlock"}]');

      const inline = await readInputSource({ source: '{"query":"inline"}', readStdin });
      const fromFile = await readInputSource({ source: `@${file}`, readStdin });
      const fromStdin = await readInputSource({ source: "-", readStdin });

      expect(inline.unwrap()).toEqual({ query: "inline" });
      expect(fromFile.unwrap()).toEqual({ query: "from file" });
      expect(fromStdin.unwrap()).toEqual([{ type: "deleteBlock" }]);
    } finally {
      await cleanup();
    }
  });

  test("refuses malformed JSON, scalars, and missing files", async () => {
    const readStdin = () => Promise.resolve("");
    for (const source of ["{", "42", "@/nonexistent/args.json"]) {
      const result = await readInputSource({ source, readStdin });
      expect(result.isErr() && result.error.code).toBe("usage_error");
    }
  });
});
