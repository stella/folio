import { describe, expect, test } from "bun:test";

import { parseArgs } from "../cli";

describe("parity CLI args", () => {
  test("parses --help and -h without treating them as input documents", () => {
    for (const arg of ["--help", "-h"]) {
      const flags = parseArgs([arg]);

      expect(flags.help).toBe(true);
      expect(flags.paths).toEqual([]);
    }
  });

  test("parses an explicit JSON output path without treating it as an input document", () => {
    const flags = parseArgs(["fixture.docx", "--max-pages", "3", "--output", "/tmp/out.json"]);

    expect(flags.paths).toEqual(["fixture.docx"]);
    expect(flags.maxPages).toBe(3);
    expect(flags.outputPath).toBe("/tmp/out.json");
  });

  test("parses the explicit playground server reuse opt-in", () => {
    const flags = parseArgs(["fixture.docx", "--reuse-server"]);

    expect(flags.paths).toEqual(["fixture.docx"]);
    expect(flags.reuseServer).toBe(true);
  });

  test("requires a path after --output", () => {
    expect(() => parseArgs(["fixture.docx", "--output"])).toThrow("--output requires a file path");
    expect(() => parseArgs(["fixture.docx", "--output", "--json"])).toThrow(
      "--output requires a file path",
    );
  });
});
