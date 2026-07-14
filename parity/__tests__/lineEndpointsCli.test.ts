import { describe, expect, test } from "bun:test";

import { parseLineEndpointCliArgs } from "../lineEndpointsCli";

describe("Word line-endpoint CLI", () => {
  test("parses capture with an explicit output", () => {
    expect(
      parseLineEndpointCliArgs([
        "capture",
        "fixture.docx",
        "--output",
        "fixture.word-lines.json",
        "--refresh-word",
      ]),
    ).toEqual({
      type: "capture",
      docxPath: "fixture.docx",
      outputPath: "fixture.word-lines.json",
      refreshWord: true,
    });
  });

  test("parses offline validation options", () => {
    expect(
      parseLineEndpointCliArgs([
        "validate",
        "fixture.docx",
        "--manifest",
        "fixture.word-lines.json",
        "--headed",
        "--reuse-server",
      ]),
    ).toEqual({
      type: "validate",
      docxPath: "fixture.docx",
      manifestPath: "fixture.word-lines.json",
      headed: true,
      reuseServer: true,
    });
  });

  test("requires explicit manifest paths", () => {
    expect(() => parseLineEndpointCliArgs(["capture", "fixture.docx"])).toThrow(
      "capture requires --output",
    );
    expect(() => parseLineEndpointCliArgs(["validate", "fixture.docx"])).toThrow(
      "validate requires --manifest",
    );
  });

  test("rejects unknown options", () => {
    expect(() =>
      parseLineEndpointCliArgs(["validate", "fixture.docx", "--manifest", "lines.json", "--ci"]),
    ).toThrow("Unknown validate option: --ci");
  });
});
