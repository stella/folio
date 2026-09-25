/**
 * Drives `folio mcp` over stdio with an MCP client, as a host application
 * would: list tools, read, suggest a change, and the refusals for a path
 * outside the allowed root and for a stale fileVersion.
 */

import { afterAll, beforeAll, describe, expect, test } from "bun:test";
import { mkdir, readFile } from "node:fs/promises";
import path from "node:path";

import { Client } from "@modelcontextprotocol/client";
import { StdioClientTransport } from "@modelcontextprotocol/client/stdio";

import { CONTRACT_PARAGRAPHS, makeTempDir, writeDocx } from "./__tests__/fixtures";
import { ISOLATED_GIT_ENV } from "./__tests__/io";
import { fileVersionOf } from "./document";
import { FOLIO_FILE_TOOLS } from "./registry";

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

type Envelope = { ok: boolean; data?: Record<string, unknown>; error?: Record<string, unknown> };

const call = async (name: string, args: Record<string, unknown>): Promise<Envelope> => {
  const result = await client.callTool({ name, arguments: args });
  const content = Array.isArray(result.content) ? result.content : [];
  const first: unknown = content.at(0);
  if (typeof first !== "object" || first === null || !("text" in first)) {
    throw new Error("tool returned no text content");
  }
  const envelope: unknown = JSON.parse(String(first.text));
  if (typeof envelope !== "object" || envelope === null || !("ok" in envelope)) {
    throw new Error("tool returned no envelope");
  }
  expect(result.isError === true).toBe(envelope.ok !== true);
  return envelope as Envelope;
};

describe("folio mcp", () => {
  test(
    "lists every registry tool with the file envelope in its schema",
    async () => {
      const { tools } = await client.listTools();

      expect(tools.map(({ name }) => name).toSorted()).toEqual(
        FOLIO_FILE_TOOLS.map(({ name }) => name).toSorted(),
      );
      const suggest = tools.find(({ name }) => name === "suggest_changes");
      expect(suggest?.inputSchema.required).toEqual(["path", "fileVersion", "operations"]);
      expect(Object.keys(suggest?.inputSchema.properties ?? {})).toContain("destination");
      const read = tools.find(({ name }) => name === "read_document");
      expect(read?.annotations?.readOnlyHint).toBe(true);
    },
    PROCESS_TEST_TIMEOUT_MS,
  );

  test(
    "reads, then suggests against the version it read",
    async () => {
      const file = await writeDocx(root, "contract.docx", CONTRACT_PARAGRAPHS);

      const read = await call("read_document", { path: "contract.docx", maxBlocks: 2 });
      expect(read.ok).toBe(true);
      const fileVersion = read.data?.["fileVersion"];
      expect(fileVersion).toBe(fileVersionOf(new Uint8Array(await readFile(file))));
      expect(JSON.stringify(read.data?.["result"])).toContain('"nextCursor"');

      const suggested = await call("suggest_changes", {
        path: file,
        fileVersion,
        operations: [{ type: "replaceInBlock", blockId: "10000002", find: "$50", replace: "$500" }],
      });
      expect(suggested.ok).toBe(true);
      expect(suggested.data?.["saveStrategy"]).toBe("selective");
      expect(suggested.data?.["author"]).toBe("MCP Reviewer");

      const stale = await call("suggest_changes", {
        path: file,
        fileVersion,
        operations: [{ type: "deleteBlock", blockId: "10000004" }],
      });
      expect(stale.error?.["code"]).toBe("stale_version");

      const changes = await call("read_changes", { path: file });
      expect(Array.isArray(changes.data?.["result"]) && changes.data["result"].length).toBe(2);
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
      expect(unversioned.error?.["code"]).toBe("invalid_input");
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
