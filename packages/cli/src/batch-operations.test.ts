/**
 * The batch operations `suggest_changes` takes beside the contract ones,
 * driven through an in-process MCP client: a tracked rename across body and
 * table with a comment in the same batch, and the refusals that write nothing.
 */

import { afterAll, beforeAll, describe, expect, test } from "bun:test";
import JSZip from "jszip";
import { mkdir, readFile } from "node:fs/promises";
import path from "node:path";

import { Client, InMemoryTransport } from "@modelcontextprotocol/client";

import { makeTempDir, writeDocx, type FixtureBlock } from "./__tests__/fixtures";
import { fileVersionOf } from "./document";
import { blockLines, createFolioMcpServer } from "./mcp";

let cleanup: () => Promise<void> = () => Promise.resolve();
let root = "";
let client: Client;

beforeAll(async () => {
  let dir = "";
  ({ dir, cleanup } = await makeTempDir());
  root = path.join(dir, "root");
  await mkdir(root);
  const server = createFolioMcpServer({
    roots: [root],
    author: "Reviewer",
    now: () => new Date("2026-01-02T03:04:05Z"),
  });
  const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
  client = new Client({ name: "folio-batch", version: "0.0.0" });
  await Promise.all([server.connect(serverTransport), client.connect(clientTransport)]);
});

afterAll(async () => {
  await client.close();
  await cleanup();
});

const BLOCKS: FixtureBlock[] = [
  {
    text: "The Supplier and the Supplier agree.",
    runs: [
      { text: "The Supplier and the " },
      { text: "Supplier", bold: true },
      { text: " agree." },
    ],
    paraId: "10000001",
  },
  {
    rows: [
      [[{ text: "Role", paraId: "10000002" }], [{ text: "Party", paraId: "10000003" }]],
      [[{ text: "Supplier", paraId: "10000004" }], [{ text: "Acme Ltd", paraId: "10000005" }]],
    ],
  },
  { text: "7.2 The Supplier's liability is capped.", paraId: "10000006" },
  { text: "7.3 Notices go to the Supplier.", paraId: "10000007" },
];

type Payload = { isError: boolean; body: Record<string, unknown> };

const call = async (name: string, args: Record<string, unknown>): Promise<Payload> => {
  const result = await client.callTool({ name, arguments: args });
  const first: unknown = Array.isArray(result.content) ? result.content.at(0) : undefined;
  if (typeof first !== "object" || first === null || !("text" in first)) {
    throw new Error(`${name} returned no text`);
  }
  return {
    isError: result.isError === true,
    body: JSON.parse(String(first.text)) as Record<string, unknown>,
  };
};

const fixture = async (name: string): Promise<{ file: string; fileVersion: string }> => {
  const file = await writeDocx(root, name, BLOCKS);
  return { file, fileVersion: fileVersionOf(new Uint8Array(await readFile(file))) };
};

const documentXml = async (file: string): Promise<string> => {
  const zip = await JSZip.loadAsync(await readFile(file));
  return (await zip.file("word/document.xml")?.async("string")) ?? "";
};

describe("read_document lines", () => {
  test("show list labels, and hide only an unnumbered heading's style id", () => {
    const lines = blockLines([
      { blockId: "1", headingLevel: 1, displayLabel: "Heading1", text: "Terms" },
      { blockId: "2", headingLevel: 2, displayLabel: "Article 1", text: "Scope" },
      { blockId: "3", displayLabel: "iii.", text: "third" },
      { blockId: "4", displayLabel: "vii)", text: "seventh" },
      { blockId: "5", displayLabel: "Article", text: "plain" },
    ]);

    expect(lines.split("\n")).toEqual([
      "[1] (h1) Terms",
      "[2] (h2) Article 1 Scope",
      "[3] iii. third",
      "[4] vii) seventh",
      "[5] Article plain",
    ]);
  });
});

describe("suggest_changes batch operations", () => {
  test("read_document puts a table row on one line", async () => {
    const { file } = await fixture("read.docx");
    const read = await call("read_document", { path: file });

    expect(String(read.body["blocks"]).split("\n")).toEqual([
      "[10000001] The Supplier and the Supplier agree.",
      "| [10000002] Role | [10000003] Party |",
      "| [10000004] Supplier | [10000005] Acme Ltd |",
      "[10000006] 7.2 The Supplier's liability is capped.",
      "[10000007] 7.3 Notices go to the Supplier.",
    ]);
  });

  test("renames everywhere and comments by quote in one tracked batch", async () => {
    const { file, fileVersion } = await fixture("rename.docx");
    const written = await call("suggest_changes", {
      path: file,
      fileVersion,
      operations: [
        { type: "replaceAll", find: "Supplier", replace: "Provider", matchCase: true },
        { type: "addComment", quote: "liability is capped", comment: "Cap too low?" },
      ],
    });

    expect(written.isError).toBe(false);
    expect(written.body).toMatchObject({
      author: "Reviewer",
      applied: 6,
      replaced: [{ find: "Supplier", count: 5 }],
    });
    expect(written.body["commentIds"]).toHaveLength(1);
    expect(written.body["fileVersion"]).toBe(fileVersionOf(new Uint8Array(await readFile(file))));

    const xml = await documentXml(file);
    expect(xml.match(/<w:ins /gu)).toHaveLength(5);
    expect(xml.match(/<w:del /gu)).toHaveLength(5);
    expect(xml).toContain('w:author="Reviewer"');
    // The bold occurrence is replaced by bold text.
    expect(xml).toMatch(/<w:ins [^>]*><w:r><w:rPr><w:b\/><\/w:rPr><w:t>Provider<\/w:t>/u);
    // The table cell is renamed too.
    expect(xml).toMatch(/<w:tc><w:p w14:paraId="10000004"[^>]*><w:del /u);

    const comments = await call("read_comments", { path: file });
    expect(comments.body["result"]).toMatchObject([
      { author: "Reviewer", text: "Cap too low?", blockId: "10000006" },
    ]);
  });

  test("direct mode replaces without tracking", async () => {
    const { file, fileVersion } = await fixture("direct.docx");
    const written = await call("suggest_changes", {
      path: file,
      fileVersion,
      mode: "direct",
      operations: [{ type: "replaceAll", find: "supplier", replace: "Provider", wholeWord: true }],
    });

    expect(written.body).toMatchObject({ applied: 5 });
    const read = await call("read_document", { path: file });
    expect(String(read.body["blocks"])).not.toContain("Supplier");
    expect(await documentXml(file)).not.toContain("<w:ins ");
  });

  test("comments on a block by id", async () => {
    const { file, fileVersion } = await fixture("block.docx");
    const written = await call("suggest_changes", {
      path: file,
      fileVersion,
      operations: [{ type: "addComment", blockId: "10000007", comment: "Which address?" }],
    });

    expect(written.body["commentIds"]).toHaveLength(1);
    const comments = await call("read_comments", { path: file });
    expect(comments.body["result"]).toMatchObject([{ blockId: "10000007" }]);
  });

  test("refuses a quote in several blocks, a missing match, and stray fields, writing nothing", async () => {
    const { file, fileVersion } = await fixture("refusals.docx");
    const refuse = async (operation: Record<string, unknown>): Promise<Record<string, unknown>> => {
      const result = await call("suggest_changes", {
        path: file,
        fileVersion,
        operations: [operation],
      });
      expect(result.isError).toBe(true);
      return result.body["error"] as Record<string, unknown>;
    };

    expect(await refuse({ type: "addComment", quote: "Supplier", comment: "x" })).toMatchObject({
      code: "ambiguous_target",
    });
    expect(await refuse({ type: "replaceAll", find: "Vendor", replace: "Provider" })).toMatchObject(
      { code: "not_found" },
    );
    expect(
      await refuse({ type: "replaceAll", find: "Supplier", replace: "Provider", comment: "x" }),
    ).toMatchObject({ code: "invalid_input" });
    expect(fileVersionOf(new Uint8Array(await readFile(file)))).toBe(fileVersion);
  });
});
