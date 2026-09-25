import { afterEach, beforeEach, describe, expect, test } from "bun:test";
import { writeFile } from "node:fs/promises";

import { buildDocx, CONTRACT_PARAGRAPHS, makeTempDir, writeDocx } from "./__tests__/fixtures";
import {
  CLI_READ_BOUNDS,
  executeReadTool,
  type FileToolCall,
  type FolioReadBounds,
} from "./execute-read";
import { findFileTool } from "./registry";

let dir = "";
let cleanup: () => Promise<void> = () => Promise.resolve();

beforeEach(async () => {
  ({ dir, cleanup } = await makeTempDir());
});

afterEach(async () => {
  await cleanup();
});

const read = async (toolName: string, call: FileToolCall, bounds = CLI_READ_BOUNDS) => {
  const tool = findFileTool(toolName);
  if (!tool) throw new Error(`${toolName} is not registered`);
  return await executeReadTool(tool, call, bounds);
};

const resultOf = (value: unknown): Record<string, unknown> => {
  if (typeof value !== "object" || value === null) throw new Error("expected an object");
  return { ...value };
};

describe("read_document", () => {
  test("labels package paraIds and synthetic ids", async () => {
    const withIds = await writeDocx(dir, "ids.docx", CONTRACT_PARAGRAPHS);
    const withoutIds = await writeDocx(
      dir,
      "plain.docx",
      CONTRACT_PARAGRAPHS.map(({ text }) => ({ text })),
    );

    const labelled = await read("read_document", { path: withIds, args: {} });
    const synthetic = await read("read_document", { path: withoutIds, args: {} });

    const sources = (data: typeof labelled) => {
      const blocks = resultOf(data.unwrap().result)["blocks"];
      if (!Array.isArray(blocks)) throw new Error("no blocks");
      return blocks.map((block) => resultOf(block)["blockIdSource"]);
    };
    expect(sources(labelled)).toEqual(["package", "package", "package", "package"]);
    expect(sources(synthetic)).toEqual(["synthetic", "synthetic", "synthetic", "synthetic"]);
  });

  test("pages with a cursor bound to the file version", async () => {
    const file = await writeDocx(dir, "contract.docx", CONTRACT_PARAGRAPHS);

    const first = (await read("read_document", { path: file, args: { maxBlocks: 3 } })).unwrap();
    const firstPage = resultOf(first.result);
    const second = await read("read_document", {
      path: file,
      args: { maxBlocks: 3, cursor: firstPage["nextCursor"] },
    });
    const secondPage = resultOf(second.unwrap().result);

    expect(firstPage["truncated"]).toBe(true);
    expect(secondPage["truncated"]).toBe(false);
    expect(secondPage["nextCursor"]).toBeUndefined();
    const ids = [firstPage, secondPage].flatMap((page) => {
      const blocks = page["blocks"];
      return Array.isArray(blocks) ? blocks.map((block) => resultOf(block)["blockId"]) : [];
    });
    expect(ids).toEqual(CONTRACT_PARAGRAPHS.map(({ paraId }) => paraId));

    await writeFile(file, await buildDocx([{ text: "Changed.", paraId: "20000001" }]));
    const stale = await read("read_document", {
      path: file,
      args: { cursor: firstPage["nextCursor"] },
    });
    expect(stale.isErr() && stale.error.code).toBe("stale_version");
  });

  test("stops a page at the response byte limit", async () => {
    const file = await writeDocx(dir, "contract.docx", CONTRACT_PARAGRAPHS);
    const whole = (await read("read_document", { path: file, args: {} })).unwrap();
    const bounds: FolioReadBounds = {
      ...CLI_READ_BOUNDS,
      maxResponseBytes: JSON.stringify(whole).length,
    };

    const page = resultOf(
      (await read("read_document", { path: file, args: {} }, bounds)).unwrap().result,
    );

    const blocks = page["blocks"];
    expect(Array.isArray(blocks) && blocks.length).toBeGreaterThan(0);
    expect(Array.isArray(blocks) && blocks.length).toBeLessThan(CONTRACT_PARAGRAPHS.length);
    expect(typeof page["nextCursor"]).toBe("string");
  });
});

describe("file envelope", () => {
  test("refuses a stale expected version and reports the current one", async () => {
    const file = await writeDocx(dir, "contract.docx", CONTRACT_PARAGRAPHS);
    const current = (await read("read_changes", { path: file, args: {} })).unwrap().fileVersion;

    const stale = await read("read_changes", { path: file, fileVersion: "0".repeat(64), args: {} });

    expect(stale.isErr() && stale.error.code).toBe("stale_version");
    expect(stale.isErr() && stale.error.details).toEqual({
      expected: "0".repeat(64),
      actual: current,
    });
  });

  test("refuses missing files and non-packages", async () => {
    const text = `${dir}/notes.docx`;
    await writeFile(text, "not a zip");

    const missing = await read("read_document", { path: `${dir}/missing.docx`, args: {} });
    const invalid = await read("read_document", { path: text, args: {} });

    expect(missing.isErr() && missing.error.code).toBe("not_found");
    expect(invalid.isErr() && invalid.error.code).toBe("invalid_document");
  });

  test("caps find_text matches while counting all of them", async () => {
    const file = await writeDocx(dir, "contract.docx", CONTRACT_PARAGRAPHS);
    const bounds: FolioReadBounds = { ...CLI_READ_BOUNDS, maxMatches: 1 };

    const found = resultOf(
      (await read("find_text", { path: file, args: { query: "pay" } }, bounds)).unwrap().result,
    );

    expect(Array.isArray(found["matches"]) && found["matches"].length).toBe(1);
    expect(found["totalMatches"]).toBe(3);
    expect(found["truncated"]).toBe(true);
  });

  test("passes tool argument errors through as invalid input", async () => {
    const file = await writeDocx(dir, "contract.docx", CONTRACT_PARAGRAPHS);

    const result = await read("find_text", { path: file, args: {} });

    expect(result.isErr() && result.error.code).toBe("invalid_input");
  });
});
