/**
 * Drives `folio mcp` over stdio with an MCP client, as a host application
 * would: list tools, read, suggest a change, reach the unlisted tools through
 * the capability tools, and the refusals for a path outside the allowed root
 * and for a stale fileVersion.
 */

import { afterAll, beforeAll, describe, expect, test } from "bun:test";
import { mkdir, readFile, stat, writeFile } from "node:fs/promises";
import path from "node:path";

import { Client } from "@modelcontextprotocol/client";
import { StdioClientTransport } from "@modelcontextprotocol/client/stdio";

import { buildDocx, CONTRACT_PARAGRAPHS, makeTempDir, writeDocx } from "./__tests__/fixtures";
import { ISOLATED_GIT_ENV } from "./__tests__/io";
import { MALFORMED_PACKAGES, TOOL_ARGUMENTS } from "./__tests__/malformed-packages";
import { fileVersionOf } from "./document";
import { listMcpTools } from "./mcp";
import { FOLIO_FILE_TOOLS, toolAccess } from "./registry";

const BIN = path.join(import.meta.dir, "bin.ts");

/** The server is a separate process; a loaded machine needs more than the default 5 s. */
const PROCESS_TEST_TIMEOUT_MS = 120_000;

let dir = "";
let root = "";
let cleanup: () => Promise<void> = () => Promise.resolve();
let client: Client;

beforeAll(async () => {
  ({ dir, cleanup } = await makeTempDir());
  root = path.join(dir, "root");
  await mkdir(root);
  const env: Record<string, string> = { FOLIO_AUTHOR: "MCP Reviewer" };
  for (const [key, value] of Object.entries(ISOLATED_GIT_ENV)) {
    if (value !== undefined) env[key] = value;
  }
  client = new Client({ name: "folio-test", version: "0.0.0" });
  await client.connect(
    new StdioClientTransport({
      command: process.execPath,
      args: [BIN, "mcp", "--root", root],
      env,
      stderr: "ignore",
    }),
  );
}, PROCESS_TEST_TIMEOUT_MS);

afterAll(async () => {
  await client.close();
  await cleanup();
});

const panicMissing = (message: string): never => {
  throw new Error(message);
};

type Envelope = { ok: boolean; data?: Record<string, unknown>; error?: Record<string, unknown> };

const call = async (name: string, args: Record<string, unknown>): Promise<Envelope> => {
  const result = await client.callTool({ name, arguments: args });
  const content = Array.isArray(result.content) ? result.content : [];
  const first: unknown = content.at(0);
  if (typeof first !== "object" || first === null || !("text" in first)) {
    throw new Error("tool returned no text content");
  }
  const payload: unknown = JSON.parse(String(first.text));
  if (typeof payload !== "object" || payload === null) {
    throw new Error("tool returned no JSON object");
  }
  const failed = result.isError === true;
  expect("error" in payload).toBe(failed);
  return failed
    ? { ok: false, error: (payload as { error: Record<string, unknown> }).error }
    : { ok: true, data: payload as Record<string, unknown> };
};

const LISTED = [
  "read_document",
  "find_text",
  "read_comments",
  "read_changes",
  "suggest_changes",
  "add_comment",
  "list_capabilities",
  "describe_capability",
  "invoke_capability",
];

const versionOf = async (file: string): Promise<string> =>
  fileVersionOf(new Uint8Array(await readFile(file)));

describe("folio mcp", () => {
  test(
    "lists the frequent tools compactly and the rest behind the capability tools",
    async () => {
      const { tools } = await client.listTools();

      expect(tools.map(({ name }) => name)).toEqual(LISTED);
      expect(Buffer.byteLength(JSON.stringify(tools))).toBeLessThan(8 * 1024);
      expect(JSON.stringify(tools)).not.toContain("destination");
      const suggest = tools.find(({ name }) => name === "suggest_changes");
      expect(suggest?.inputSchema.required).toEqual(["path", "fileVersion", "operations"]);
      expect(suggest?.annotations?.destructiveHint).toBe(true);
      const read = tools.find(({ name }) => name === "read_document");
      expect(read?.annotations?.readOnlyHint).toBe(true);
      expect(read?.annotations?.destructiveHint).toBe(false);

      const listed = await call("list_capabilities", {});
      const items = (listed.data?.["items"] ?? []) as { id: string }[];
      const ids = items.map(({ id }) => id);
      expect([...LISTED, ...ids].toSorted()).toEqual(
        [...FOLIO_FILE_TOOLS.map(({ name }) => name), ...LISTED.slice(-3)].toSorted(),
      );
      const outlined = await call("describe_capability", { capability: "suggest_changes" });
      expect(outlined.data?.["parameters"]).toMatchObject({ destination: "string" });
      expect(outlined.data?.["description"]).toContain("replaceAll");
      const described = await call("describe_capability", {
        capability: "suggest_changes",
        detail: "full",
      });
      const schema = described.data?.["inputSchema"] as { properties: Record<string, unknown> };
      expect(Object.keys(schema.properties)).toContain("destination");
      expect(described.data?.["description"]).toContain("replaceInBlock");
      expect(listMcpTools()).toEqual(tools);
    },
    PROCESS_TEST_TIMEOUT_MS,
  );

  test(
    "chains a tracked edit and a comment on the same paragraph without re-reading",
    async () => {
      const file = await writeDocx(root, "unidentified.docx", [
        { text: "The Supplier shall deliver." },
        { text: "The Supplier is liable up to $50." },
      ]);

      const read = await call("read_document", { path: file });
      const lines = String(read.data?.["blocks"]).split("\n");
      expect(lines).toHaveLength(2);
      const [first, second] = lines.map(
        (line) => /^\[(?<id>[0-9A-F]{8})\] /u.exec(line)?.groups?.["id"],
      );
      expect(lines[1]).toBe(`[${String(second)}] The Supplier is liable up to $50.`);

      const found = await call("find_text", { path: file, query: "Supplier", matchCase: "true" });
      const matches = found.data?.["matches"] as { range: Record<string, unknown> }[];
      expect(matches).toHaveLength(2);
      expect(matches[0]?.range).not.toHaveProperty("type");

      const suggested = await call("suggest_changes", {
        path: file,
        fileVersion: read.data?.["fileVersion"],
        operations: [
          { type: "replaceInBlock", blockId: first, find: "Supplier", replace: "Provider" },
          { type: "replaceRange", range: matches[1]?.range, replace: "Provider" },
        ],
      });
      expect(suggested.data).toEqual({
        fileVersion: await versionOf(file),
        author: "MCP Reviewer",
        applied: 2,
      });

      const commented = await call("add_comment", {
        path: file,
        fileVersion: suggested.data?.["fileVersion"],
        blockId: second,
        text: "Cap seems low.",
      });
      expect(commented.ok).toBe(true);
      expect(typeof commented.data?.["commentId"]).toBe("string");

      const changes = await call("read_changes", { path: file });
      expect(String(changes.data?.["changes"]).split("\n")).toHaveLength(4);
    },
    PROCESS_TEST_TIMEOUT_MS,
  );

  test(
    "runs an unlisted tool through invoke_capability",
    async () => {
      const file = await writeDocx(root, "resolve.docx", CONTRACT_PARAGRAPHS);
      await call("suggest_changes", {
        path: file,
        fileVersion: await versionOf(file),
        operations: [{ type: "replaceInBlock", blockId: "10000002", find: "$50", replace: "$5" }],
      });
      const current = await versionOf(file);

      const checked = await call("invoke_capability", {
        capability: "resolve_changes",
        input: { path: file, fileVersion: current, action: "accept", all: true },
        validate_only: true,
      });
      const unchanged = await versionOf(file);
      const accepted = await call("invoke_capability", {
        capability: "resolve_changes",
        input: { path: file, fileVersion: current, action: "Accept", all: "yes" },
      });

      expect(checked.data?.["valid"]).toBe(true);
      expect(unchanged).toBe(current);
      expect(accepted.data).toMatchObject({ resolved: 2, remaining: 0 });
    },
    PROCESS_TEST_TIMEOUT_MS,
  );

  test(
    "reads, then suggests against the version it read",
    async () => {
      const file = await writeDocx(root, "contract.docx", CONTRACT_PARAGRAPHS);

      const read = await call("read_document", { path: "contract.docx", maxBlocks: "2" });
      expect(read.ok).toBe(true);
      const fileVersion = read.data?.["fileVersion"];
      expect(fileVersion).toBe(await versionOf(file));
      expect(read.data?.["blocks"]).toBe(
        "[10000001] (h1) Payment\n[10000002] The buyer pays $50 on signing.",
      );
      expect(read.data?.["nextCursor"]).toBeString();

      const suggested = await call("suggest_changes", {
        path: file,
        fileVersion,
        operations: [{ type: "replaceInBlock", blockId: "10000002", find: "$50", replace: "$500" }],
      });
      expect(suggested.ok).toBe(true);
      expect(suggested.data?.["applied"]).toBe(1);

      const stale = await call("suggest_changes", {
        path: file,
        fileVersion,
        operations: [{ type: "deleteBlock", blockId: "10000004" }],
      });
      expect(stale.error).toMatchObject({ code: "stale_version", retryable: true });

      const changes = await call("read_changes", { path: file });
      const lines = String(changes.data?.["changes"]).split("\n");
      expect(lines).toHaveLength(2);
      expect(
        lines.every(
          (line) => line.endsWith('"$50" MCP Reviewer') || line.endsWith('"$500" MCP Reviewer'),
        ),
      ).toBe(true);
    },
    PROCESS_TEST_TIMEOUT_MS,
  );

  test(
    "replaces an existing destination only with overwrite and its version, keeping a backup",
    async () => {
      const source = await writeDocx(root, "draft.docx", CONTRACT_PARAGRAPHS);
      const target = await writeDocx(root, "target.docx", [{ text: "Keep.", paraId: "20000001" }]);
      await writeFile(path.join(root, "package.json"), "{}");
      const fileVersion = fileVersionOf(new Uint8Array(await readFile(source)));
      const targetVersion = fileVersionOf(new Uint8Array(await readFile(target)));
      const comment = { path: source, fileVersion, blockId: "10000002", text: "?" };

      const refused = await call("add_comment", {
        ...comment,
        destination: target,
        overwrite: true,
      });
      const notDocx = await call("add_comment", {
        ...comment,
        destination: "package.json",
        overwrite: true,
      });
      const compareUnversioned = await call("compare_documents", {
        path: source,
        revisedPath: target,
        destination: "redline.docx",
      });
      expect(refused.error?.["code"]).toBe("invalid_input");
      expect(notDocx.error?.["code"]).toBe("invalid_destination");
      expect(compareUnversioned.error?.["code"]).toBe("invalid_input");
      expect(fileVersionOf(new Uint8Array(await readFile(target)))).toBe(targetVersion);
      expect(await readFile(path.join(root, "package.json"), "utf8")).toBe("{}");

      const replaced = await call("add_comment", {
        ...comment,
        destination: target,
        overwrite: true,
        expectedDestinationVersion: targetVersion,
      });
      expect(replaced.ok).toBe(true);
      expect(replaced.data?.["path"]).toBe(target);
      const backup = path.join(root, ".folio", "backups", "target.docx", `${targetVersion}.docx`);
      expect((await stat(backup)).isFile()).toBe(true);
    },
    PROCESS_TEST_TIMEOUT_MS,
  );

  test(
    "refuses paths outside the allowed root and changes without a fileVersion",
    async () => {
      const outside = await writeDocx(dir, "outside.docx", CONTRACT_PARAGRAPHS);
      const inside = await writeDocx(root, "inside.docx", CONTRACT_PARAGRAPHS);

      const read = await call("read_document", { path: outside });
      const escape = await call("read_document", { path: "../outside.docx" });
      const destination = await call("add_comment", {
        path: inside,
        fileVersion: fileVersionOf(new Uint8Array(await readFile(inside))),
        destination: path.join(dir, "copy.docx"),
        blockId: "10000002",
        text: "?",
      });
      const unversioned = await call("add_comment", {
        path: inside,
        blockId: "10000002",
        text: "?",
      });

      expect(read.error?.["code"]).toBe("outside_root");
      expect(escape.error?.["code"]).toBe("outside_root");
      expect(destination.error?.["code"]).toBe("outside_root");
      expect(unversioned.error?.["code"]).toBe("validation_error");
    },
    PROCESS_TEST_TIMEOUT_MS,
  );

  test(
    "every tool refuses a file that is not a WordprocessingML package",
    async () => {
      const valid = await buildDocx(CONTRACT_PARAGRAPHS);
      const other = path.join(root, "malformed-other.docx");
      await writeFile(other, valid);
      const outcomes: string[] = [];
      const expected: string[] = [];
      for (const [index, { name, build }] of MALFORMED_PACKAGES.entries()) {
        const bytes = await build(valid);
        const file = path.join(root, `malformed-${String(index)}.docx`);
        await writeFile(file, bytes);
        const fileVersion = fileVersionOf(bytes);
        for (const tool of FOLIO_FILE_TOOLS) {
          const calls: [string, Record<string, unknown>][] =
            tool.type === "compare"
              ? [
                  ["base", { path: file, revisedPath: other }],
                  ["revised", { path: other, revisedPath: file }],
                  [
                    "redline",
                    {
                      path: file,
                      fileVersion,
                      revisedPath: other,
                      destination: path.join(root, `malformed-${String(index)}-redline.docx`),
                    },
                  ],
                ]
              : [
                  [
                    toolAccess(tool),
                    {
                      path: file,
                      fileVersion,
                      ...(TOOL_ARGUMENTS[tool.name] ??
                        panicMissing(`no arguments for ${tool.name}`)),
                      ...(tool.type === "resolveChanges" && { action: "accept" }),
                    },
                  ],
                ];
          for (const [variant, args] of calls) {
            const label = `${name} / ${tool.name} (${variant})`;
            const envelope = await call(tool.name, args);
            outcomes.push(`${label}: ${String(envelope.error?.["code"])}`);
            expected.push(`${label}: invalid_document`);
          }
        }
        expect(new Uint8Array(await readFile(file))).toEqual(bytes);
      }
      expect(outcomes).toEqual(expected);
      expect(new Uint8Array(await readFile(other))).toEqual(valid);
    },
    PROCESS_TEST_TIMEOUT_MS,
  );

  test(
    "serves the operation schema and the about page as resources",
    async () => {
      const { resources } = await client.listResources();
      const schema = await client.readResource({ uri: "folio://schema/operations" });

      expect(resources.map(({ uri }) => uri)).toEqual([
        "folio://about",
        "folio://schema/operations",
      ]);
      const first = schema.contents.at(0);
      expect(first !== undefined && "text" in first && first.text).toContain('"operations"');
    },
    PROCESS_TEST_TIMEOUT_MS,
  );
});
