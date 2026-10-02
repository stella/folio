/**
 * What a model pays for the MCP surface, in characters: the tool listing and
 * instructions it reads on every turn, a description of suggest_changes, and
 * a read of a four-page contract. Printed for comparison across changes and
 * held under budgets.
 */

import { afterAll, beforeAll, expect, test } from "bun:test";
import { mkdir } from "node:fs/promises";
import path from "node:path";

import { Client, InMemoryTransport } from "@modelcontextprotocol/client";

import { fourPageContract, makeTempDir, writeDocx } from "./__tests__/fixtures";
import { createFolioMcpServer } from "./mcp";

let cleanup: () => Promise<void> = () => Promise.resolve();
let root = "";
let client: Client;

beforeAll(async () => {
  let dir = "";
  ({ dir, cleanup } = await makeTempDir());
  root = path.join(dir, "root");
  await mkdir(root);
  const server = createFolioMcpServer({ roots: [root], author: "Reviewer" });
  const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
  client = new Client({ name: "folio-size", version: "0.0.0" });
  await Promise.all([server.connect(serverTransport), client.connect(clientTransport)]);
});

afterAll(async () => {
  await client.close();
  await cleanup();
});

const textOf = async (name: string, args: Record<string, unknown>): Promise<string> => {
  const result = await client.callTool({ name, arguments: args });
  const first: unknown = Array.isArray(result.content) ? result.content.at(0) : undefined;
  if (typeof first !== "object" || first === null || !("text" in first)) {
    throw new Error(`${name} returned no text`);
  }
  expect(result.isError).not.toBe(true);
  return String(first.text);
};

test("the MCP surface stays small", async () => {
  const file = await writeDocx(root, "contract.docx", fourPageContract());
  const sizes = {
    toolsList: JSON.stringify((await client.listTools()).tools).length,
    instructions: (client.getInstructions() ?? "").length,
    describeSuggestChanges: (await textOf("describe_capability", { capability: "suggest_changes" }))
      .length,
    describeSuggestChangesFull: (
      await textOf("describe_capability", { capability: "suggest_changes", detail: "full" })
    ).length,
    readFourPageContract: (await textOf("read_document", { path: file })).length,
  };
  console.log(JSON.stringify(sizes));

  expect(sizes.toolsList).toBeLessThan(6_000);
  expect(sizes.instructions).toBeLessThan(800);
  expect(sizes.describeSuggestChanges).toBeLessThan(3_000);
  expect(sizes.readFourPageContract).toBeLessThan(14_000);
});
