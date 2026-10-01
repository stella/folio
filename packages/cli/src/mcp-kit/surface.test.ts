import { describe, expect, test } from "bun:test";

import {
  advertisedBytes,
  CAPABILITY_TOOL_NAMES,
  compactSchema,
  createToolSurface,
  failure,
  hoistRepeatedSchemas,
  success,
  type ToolCallResult,
  type ToolDefinition,
} from "./index";

type Context = { calls: { name: string; args: Record<string, unknown> }[] };

const record = (name: string) => (args: Record<string, unknown>, context: Context) => {
  context.calls.push({ name, args });
  return Promise.resolve(success({ tool: name, args }));
};

const SEARCH: ToolDefinition<Context> = {
  name: "search",
  summary: "Search.",
  access: "read",
  inputSchema: {
    type: "object",
    properties: {
      query: { type: "string", description: "What to look for, in detail." },
      exact: { type: "boolean", description: "Match case." },
      limit: { type: "integer", minimum: 1, maximum: 50 },
    },
    required: ["query"],
  },
  direct: {
    inputSchema: {
      type: "object",
      properties: { query: { type: "string", description: "Text." } },
      required: ["query"],
    },
  },
  run: record("search"),
};

const ARCHIVE: ToolDefinition<Context> = {
  name: "archive_item",
  summary: "Archive an item.",
  guide: "Archived items can be restored within 30 days.",
  access: "write",
  domain: "items",
  inputSchema: {
    type: "object",
    properties: {
      id: { type: "string" },
      mode: { type: "string", enum: ["soft", "hard"] },
      force: { type: "boolean" },
    },
    required: ["id"],
  },
  exactProperties: ["force"],
  run: record("archive_item"),
};

const STATS: ToolDefinition<Context> = {
  name: "item_stats",
  summary: "Count items.",
  access: "read",
  domain: "items",
  inputSchema: { type: "object", properties: {} },
  run: () => Promise.resolve(failure({ code: "stale", message: "Out of date.", retryable: true })),
};

const THROWS: ToolDefinition<Context> = {
  name: "explode",
  summary: "Throw.",
  access: "read",
  domain: "misc",
  inputSchema: { type: "object", properties: {} },
  run: () => Promise.reject(new Error("boom")),
};

const surface = createToolSurface({ tools: [SEARCH, ARCHIVE, STATS, THROWS] });

const payload = (result: ToolCallResult): Record<string, unknown> => {
  const first = result.content.at(0);
  if (first === undefined) throw new Error("no content");
  return JSON.parse(first.text) as Record<string, unknown>;
};

const call = async (name: string, args: unknown) => {
  const context: Context = { calls: [] };
  const result = await surface.callTool(name, args, context);
  return { result, body: payload(result), calls: context.calls };
};

describe("listing", () => {
  test("lists direct tools with compact schemas, then the three capability tools", () => {
    const tools = surface.listTools();

    expect(tools.map(({ name }) => name)).toEqual([
      "search",
      CAPABILITY_TOOL_NAMES.list,
      CAPABILITY_TOOL_NAMES.describe,
      CAPABILITY_TOOL_NAMES.invoke,
    ]);
    expect(tools[0]?.inputSchema).toEqual({
      type: "object",
      properties: { query: { type: "string" } },
      required: ["query"],
    });
    expect(JSON.stringify(tools)).not.toContain('"description":"Text."');
    expect(tools[0]?.annotations).toEqual({
      readOnlyHint: true,
      destructiveHint: false,
      openWorldHint: false,
    });
    expect(tools.at(-1)?.annotations.destructiveHint).toBe(true);
    expect(advertisedBytes(tools)).toBe(new TextEncoder().encode(JSON.stringify(tools)).length);
  });

  test("a surface with no lazy tools lists no capability tools", () => {
    const only = createToolSurface({ tools: [SEARCH] });
    expect(only.listTools().map(({ name }) => name)).toEqual(["search"]);
  });

  test("refuses duplicate or reserved names", () => {
    expect(() => createToolSurface({ tools: [SEARCH, SEARCH] })).toThrow("taken");
    expect(() =>
      createToolSurface({ tools: [{ ...STATS, name: CAPABILITY_TOOL_NAMES.invoke }] }),
    ).toThrow("taken");
  });

  test("compactSchema strips annotations at every depth but keeps property names", () => {
    expect(
      compactSchema({
        type: "object",
        description: "x",
        properties: {
          description: { type: "string", description: "a property named description" },
          nested: { type: "array", items: { type: "object", title: "t", default: {} } },
        },
      }),
    ).toEqual({
      type: "object",
      properties: {
        description: { type: "string" },
        nested: { type: "array", items: { type: "object" } },
      },
    });
  });

  test("hoistRepeatedSchemas states a repeated shape once", () => {
    const shape = {
      type: "object",
      properties: { left: { type: "integer" }, right: { type: "integer" } },
    };
    const hoisted = hoistRepeatedSchemas(
      { type: "object", properties: { first: shape, second: shape, third: { type: "string" } } },
      { minBytes: 40 },
    );

    expect(hoisted).toEqual({
      type: "object",
      properties: {
        first: { $ref: "#/$defs/first" },
        second: { $ref: "#/$defs/first" },
        third: { type: "string" },
      },
      $defs: { first: shape },
    });
    expect(hoistRepeatedSchemas({ type: "object", properties: { only: shape } })).toEqual({
      type: "object",
      properties: { only: shape },
    });
  });

  test("compactSchema drops a maximum that only says integer", () => {
    expect(
      compactSchema({ type: "integer", minimum: 0, maximum: Number.MAX_SAFE_INTEGER }),
    ).toEqual({ type: "integer", minimum: 0 });
  });

  test("compactSchema keeps descriptions as many property levels deep as asked", () => {
    const schema = {
      type: "object",
      description: "the tool",
      properties: {
        ops: {
          type: "array",
          description: "the operations",
          items: {
            type: "object",
            properties: { find: { type: "string", description: "text to find" } },
          },
        },
      },
    };

    expect(compactSchema(schema, { describedDepth: 1 })).toEqual({
      type: "object",
      properties: {
        ops: {
          type: "array",
          description: "the operations",
          items: { type: "object", properties: { find: { type: "string" } } },
        },
      },
    });
  });
});

describe("capability tools", () => {
  test("list_capabilities pages the unlisted tools and filters by domain and access", async () => {
    const all = await call(CAPABILITY_TOOL_NAMES.list, {});
    const items = await call(CAPABILITY_TOOL_NAMES.list, { domain: "items", access: "write" });
    const first = await call(CAPABILITY_TOOL_NAMES.list, { limit: 1 });
    const second = await call(CAPABILITY_TOOL_NAMES.list, {
      limit: 1,
      cursor: first.body["nextCursor"],
    });

    expect(all.body).toEqual({
      items: [
        { id: "archive_item", summary: "Archive an item.", access: "write", destructive: true },
        { id: "explode", summary: "Throw.", access: "read" },
        { id: "item_stats", summary: "Count items.", access: "read" },
      ],
      nextCursor: null,
    });
    expect(items.body["items"]).toEqual([
      { id: "archive_item", summary: "Archive an item.", access: "write", destructive: true },
    ]);
    expect(second.body["items"]).toEqual([{ id: "explode", summary: "Throw.", access: "read" }]);
  });

  test("pages ids in one ordinal order, whatever their case, digits or underscores", async () => {
    const names = ["b_tool", "B2", "a10", "a9", "A_x", "item", "Item_2", "_z"];
    const mixed = createToolSurface({ tools: names.map((name) => ({ ...STATS, name })) });
    const seen: string[] = [];
    let cursor: unknown;
    for (let page = 0; page <= names.length; page += 1) {
      const result = await mixed.callTool(
        CAPABILITY_TOOL_NAMES.list,
        { limit: 3, ...(typeof cursor === "string" && { cursor }) },
        { calls: [] },
      );
      const body = payload(result);
      seen.push(...(body["items"] as { id: string }[]).map(({ id }) => id));
      cursor = body["nextCursor"];
      if (cursor === null) break;
    }

    expect(seen).toEqual(names.toSorted());
  });

  test("its own arguments are read strictly", async () => {
    const typo = await call(CAPABILITY_TOOL_NAMES.list, { domian: "items" });
    const stringy = await call(CAPABILITY_TOOL_NAMES.invoke, {
      capability: "archive_item",
      input: { id: "a" },
      validate_only: "true",
    });

    expect(typo.result.isError).toBe(true);
    expect(typo.body).toMatchObject({
      error: { code: "validation_error", issues: [{ path: "domian" }], retryable: true },
    });
    expect(stringy.body).toMatchObject({ error: { issues: [{ path: "validate_only" }] } });
    expect(stringy.calls).toEqual([]);
  });

  test("describe_capability outlines any tool, and gives its full schema on request", async () => {
    const compact = await call(CAPABILITY_TOOL_NAMES.describe, { capability: "archive_item" });
    const lazy = await call(CAPABILITY_TOOL_NAMES.describe, {
      capability: "archive_item",
      detail: "full",
    });
    const direct = await call(CAPABILITY_TOOL_NAMES.describe, {
      capability: "search",
      detail: "full",
    });
    const unknown = await call(CAPABILITY_TOOL_NAMES.describe, { capability: "archive" });
    const badDetail = await call(CAPABILITY_TOOL_NAMES.describe, {
      capability: "search",
      detail: "everything",
    });

    expect(compact.body).toEqual({
      id: "archive_item",
      description: "Archive an item.",
      access: "write",
      destructive: true,
      parameters: { id: "string (required)", mode: '"soft" | "hard"', force: "boolean" },
      more: 'detail: "full" returns the full input schema.',
    });
    expect(lazy.body).toEqual({
      id: "archive_item",
      description: "Archive an item.\nArchived items can be restored within 30 days.",
      access: "write",
      destructive: true,
      inputSchema: ARCHIVE.inputSchema,
    });
    expect(direct.body["inputSchema"]).toEqual(SEARCH.inputSchema);
    expect(badDetail.body).toMatchObject({ error: { issues: [{ path: "detail" }] } });
    expect(unknown.body).toMatchObject({
      error: { code: "not_found", retryable: false },
    });
    expect(String((unknown.body["error"] as { hint: string }).hint)).toContain('"archive_item"');
  });

  test("invoke_capability runs a tool with its input, or only validates it", async () => {
    const ran = await call(CAPABILITY_TOOL_NAMES.invoke, {
      capability: "archive_item",
      input: { id: "a", mode: "HARD" },
    });
    const checked = await call(CAPABILITY_TOOL_NAMES.invoke, {
      capability: "archive_item",
      input: { id: "a" },
      validate_only: true,
    });

    expect(ran.calls).toEqual([{ name: "archive_item", args: { id: "a", mode: "hard" } }]);
    expect(ran.result.content.at(1)?.text).toBe('Input read: Read "HARD" as "hard".');
    expect(checked.body).toEqual({ valid: true, input: { id: "a" } });
    expect(checked.calls).toEqual([]);
  });
});

describe("calling a tool", () => {
  test("reads arguments leniently and says how", async () => {
    const { body, calls, result } = await call("search", {
      query: "x",
      exact: "yes",
      limit: "5",
      unusedOptional: undefined,
    });

    expect(result.isError).toBe(false);
    expect(calls).toEqual([{ name: "search", args: { query: "x", exact: true, limit: 5 } }]);
    expect(body).toEqual({ tool: "search", args: { query: "x", exact: true, limit: 5 } });
    expect(result.content.at(1)?.text).toContain('Read "yes" as true.');
  });

  test("drops null optionals, and keeps exact properties exactly as sent", async () => {
    const { calls } = await call("archive_item", { id: "a", mode: null, force: "true" });

    expect(calls).toEqual([{ name: "archive_item", args: { id: "a", force: "true" } }]);
  });

  test("refuses inherited names such as toString and __proto__ as parameters", async () => {
    // JSON.parse makes `__proto__` an own key, as a client's arguments would.
    const args: unknown = JSON.parse('{"query":"x","toString":"y","__proto__":{"polluted":true}}');
    const tool = await call("search", args);
    const meta = await call(
      CAPABILITY_TOOL_NAMES.describe,
      JSON.parse('{"capability":"search","__proto__":{"detail":"full"},"constructor":1}'),
    );

    expect(tool.result.isError).toBe(true);
    expect(tool.body).toMatchObject({
      error: { issues: [{ path: "toString" }, { path: "__proto__" }] },
    });
    expect(tool.calls).toEqual([]);
    expect(meta.body).toMatchObject({
      error: { issues: [{ path: "__proto__" }, { path: "constructor" }] },
    });
    expect(({} as Record<string, unknown>)["polluted"]).toBeUndefined();
  });

  test("refuses unknown and missing arguments without running", async () => {
    const unknown = await call("search", { query: "x", qeury: "y" });
    const missing = await call("archive_item", {});
    const ambiguous = await call("search", { query: "x", exact: "perhaps" });

    for (const outcome of [unknown, missing, ambiguous]) {
      expect(outcome.result.isError).toBe(true);
      expect(outcome.calls).toEqual([]);
    }
    expect(unknown.body).toMatchObject({
      error: {
        code: "validation_error",
        issues: [{ path: "qeury", message: "Unknown parameter: qeury" }],
        hint: "Accepted parameters: query, exact, limit.",
      },
    });
    expect(missing.body).toMatchObject({ error: { issues: [{ path: "id" }] } });
    expect(ambiguous.body).toMatchObject({ error: { issues: [{ path: "exact" }] } });
  });

  test("a tool's own failure and a throw both come back in the one envelope", async () => {
    const stale = await call("item_stats", {});
    const thrown = await call("explode", {});

    expect(stale.body).toEqual({
      error: { code: "stale", message: "Out of date.", retryable: true },
    });
    expect(thrown.body).toEqual({
      error: { code: "internal_error", message: "boom", retryable: false },
    });
    expect(thrown.result.isError).toBe(true);
  });

  test("an unknown tool suggests the close names", async () => {
    const { body } = await call("serach", {});

    expect(body).toMatchObject({ error: { code: "unknown_tool" } });
    expect(String((body["error"] as { hint: string }).hint)).toContain('"search"');
  });
});
