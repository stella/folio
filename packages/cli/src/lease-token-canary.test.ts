import { afterEach, beforeEach, describe, expect, test } from "bun:test";
import { Client } from "@modelcontextprotocol/client";
import { InMemoryTransport } from "@modelcontextprotocol/server";
import { hostname } from "node:os";
import path from "node:path";

import { CONTRACT_PARAGRAPHS, makeTempDir, writeDocx } from "./__tests__/fixtures";
import { captureIo, dataOf, envelopeOf } from "./__tests__/io";
import { fileVersionOf } from "./document";
import { acquireEditorLease, acquireLeaseForWrite, flushRequestPathFor } from "./editor-lease";
import { createFolioMcpServer } from "./mcp";
import { FOLIO_FILE_TOOLS, toolAccess } from "./registry";
import { lockPathFor } from "./lock";
import { runFolioCli } from "./cli";
import { readFile, rm, writeFile } from "node:fs/promises";
import { randomUUID } from "node:crypto";

let dir = "";
let cleanup = async (): Promise<void> => {};

beforeEach(async () => {
  ({ dir, cleanup } = await makeTempDir());
});

afterEach(async () => {
  await cleanup();
});

const CANARY_PREFIX = "folio-lease-token-canary-";

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

const sectionHandleFrom = (text: string): Record<string, unknown> => {
  const data = envelopeOf(text)["data"];
  if (!isRecord(data) || !isRecord(data["result"])) {
    throw new Error("outline response omitted its result");
  }
  const sections = data["result"]["sections"];
  const firstSection = Array.isArray(sections) ? sections.at(0) : undefined;
  if (!isRecord(firstSection) || !isRecord(firstSection["handle"])) {
    throw new Error("outline response omitted its first section handle");
  }
  return firstSection["handle"];
};

const writeHeldLease = async (documentPath: string, token: string, acceptsFlush = false) => {
  await writeFile(
    lockPathFor(documentPath),
    JSON.stringify({
      owner: "canary-editor",
      pid: process.pid,
      host: hostname(),
      txId: "canary-test",
      token,
      acquiredAt: new Date().toISOString(),
      expiresAt: new Date(Date.now() + 60_000).toISOString(),
      ...(acceptsFlush && { acceptsFlush: true }),
    }),
  );
};

const removeHeldLease = async (documentPath: string): Promise<void> => {
  await rm(lockPathFor(documentPath), { force: true });
  await rm(`${lockPathFor(documentPath)}.swap`, { force: true });
};

const argsFor = (name: string, revisedPath: string): Record<string, unknown> => {
  switch (name) {
    case "suggest_changes":
      return {
        operations: [{ type: "replaceInBlock", blockId: "10000002", find: "$50", replace: "$500" }],
      };
    case "add_comment":
      return { blockId: "10000002", text: "Canary sweep comment." };
    case "reply_comment":
      return { commentId: "1", text: "Canary sweep reply." };
    case "resolve_comment":
      return { commentId: "1" };
    case "resolve_changes":
      return { action: "accept", all: true };
    case "compare_documents":
      return { revisedPath };
    default:
      throw new Error(`No write fixture for registered tool ${name}`);
  }
};

const readArgsFor = (
  name: string,
  sectionHandle: Record<string, unknown>,
): Record<string, unknown> => {
  switch (name) {
    case "read_section":
      return { handle: sectionHandle };
    case "read_story":
      return { handle: { type: "main" } };
    case "find_text":
      return { query: "Late" };
    default:
      return {};
  }
};

const assertNoCanary = (text: string, canary: string): void => {
  expect(text).not.toContain(canary);
};

describe("lease token output canary", () => {
  test("keeps tokens out of every CLI and MCP command/tool result", async () => {
    const canary = `${CANARY_PREFIX}${randomUUID()}`;
    const file = await writeDocx(dir, "contract.docx", CONTRACT_PARAGRAPHS);
    const revised = await writeDocx(dir, "revised.docx", CONTRACT_PARAGRAPHS);
    const bytes = await readFile(file);
    const version = fileVersionOf(new Uint8Array(bytes));
    const outlineCapture = captureIo();
    expect(await runFolioCli(["outline", file], outlineCapture.io)).toBe(0);
    const sectionHandle = sectionHandleFrom(outlineCapture.stdout());
    assertNoCanary(outlineCapture.stdout(), canary);
    const serialized = await writeDocx(dir, "editor-save.docx", [
      ...CONTRACT_PARAGRAPHS,
      { text: "Saved by the editor.", paraId: "10000005" },
    ]);

    // Read and status-like paths must remain usable while a writer lease exists.
    await writeHeldLease(file, canary);
    const read = captureIo();
    expect(await runFolioCli(["read", file], read.io)).toBe(0);
    assertNoCanary(read.stdout() + read.stderr(), canary);
    const changes = captureIo();
    expect(await runFolioCli(["changes", file], changes.io)).toBe(0);
    assertNoCanary(changes.stdout() + changes.stderr(), canary);
    const renderHelp = captureIo();
    expect(await runFolioCli(["render", "--help"], renderHelp.io)).toBe(0);
    assertNoCanary(renderHelp.stdout() + renderHelp.stderr(), canary);
    const rendered = captureIo();
    expect(
      await runFolioCli(
        ["render", file, "--out", path.join(dir, "preview.html"), "--output", "json"],
        rendered.io,
      ),
    ).toBe(0);
    expect(envelopeOf(rendered.stdout())["ok"]).toBe(true);
    assertNoCanary(rendered.stdout() + rendered.stderr(), canary);

    const readTools = FOLIO_FILE_TOOLS.filter((tool) => toolAccess(tool) === "read");
    const cliReadCommands = readTools.flatMap((tool) => tool.commands.map(({ name }) => name));
    for (const commandName of cliReadCommands) {
      const tool = readTools.find(({ commands }) =>
        commands.some(({ name }) => name === commandName),
      );
      if (tool === undefined) throw new Error(`No registry tool for command ${commandName}`);
      const captured = captureIo();
      const exitCode = await runFolioCli(
        [commandName, file, "--input", JSON.stringify(readArgsFor(tool.name, sectionHandle))],
        captured.io,
      );
      if (exitCode !== 0) {
        throw new Error(`${commandName} failed: ${captured.stdout()} ${captured.stderr()}`);
      }
      assertNoCanary(captured.stdout() + captured.stderr(), canary);
    }
    await removeHeldLease(file);

    // The standalone save command adopts the exact token, while a wrong token
    // must report a lock refusal without serializing the holder token.
    await writeHeldLease(file, canary);
    const saved = captureIo();
    expect(
      await runFolioCli(
        [
          "save",
          file,
          "--from",
          serialized,
          "--expect-version",
          version,
          "--lease-token",
          canary,
          "--author",
          "Canary Reviewer",
          "--output",
          "json",
        ],
        saved.io,
      ),
    ).toBe(0);
    expect(dataOf(saved.stdout())["status"]).toBe("committed");
    assertNoCanary(saved.stdout() + saved.stderr(), canary);
    await removeHeldLease(file);
    const currentVersion = fileVersionOf(new Uint8Array(await readFile(file)));

    await writeHeldLease(file, canary);
    const wrongToken = captureIo();
    expect(
      await runFolioCli(
        [
          "save",
          file,
          "--from",
          serialized,
          "--expect-version",
          fileVersionOf(new Uint8Array(await readFile(file))),
          "--lease-token",
          "wrong-token",
          "--author",
          "Canary Reviewer",
          "--output",
          "json",
        ],
        wrongToken.io,
      ),
    ).toBe(10);
    expect(envelopeOf(wrongToken.stdout())).toMatchObject({
      ok: false,
      error: { code: "locked" },
    });
    assertNoCanary(wrongToken.stdout() + wrongToken.stderr(), canary);
    await removeHeldLease(file);

    await writeHeldLease(file, canary);
    const saveWithoutToken = captureIo();
    expect(
      await runFolioCli(
        [
          "save",
          file,
          "--from",
          serialized,
          "--expect-version",
          currentVersion,
          "--author",
          "Canary Reviewer",
          "--output",
          "json",
        ],
        saveWithoutToken.io,
      ),
    ).toBe(10);
    expect(envelopeOf(saveWithoutToken.stdout())).toMatchObject({
      ok: false,
      error: { code: "locked" },
    });
    assertNoCanary(saveWithoutToken.stdout() + saveWithoutToken.stderr(), canary);
    await removeHeldLease(file);

    // Derive the full registry-backed CLI write surface. Compare writes to
    // their destination; all other writes target the source in place.
    const writeTools = FOLIO_FILE_TOOLS.filter((tool) => toolAccess(tool) !== "read");
    const cliCommands = writeTools.flatMap((tool) => tool.commands.map(({ name }) => name));
    for (const [index, commandName] of cliCommands.entries()) {
      const tool = writeTools.find(({ commands }) =>
        commands.some(({ name }) => name === commandName),
      );
      if (tool === undefined) throw new Error(`No registry tool for command ${commandName}`);
      const target =
        tool.type === "compare" ? path.join(dir, `cli-redline-${String(index)}.docx`) : file;
      await writeHeldLease(target, canary);
      const input = JSON.stringify(argsFor(tool.name, revised));
      const argv = [
        commandName,
        file,
        ...(tool.type === "compare" ? [revised] : []),
        "--input",
        input,
        "--expect-version",
        currentVersion,
        ...(tool.type === "compare" ? ["--out", target] : ["--in-place"]),
      ];
      const captured = captureIo();
      const exitCode = await runFolioCli(argv, captured.io);
      expect(exitCode).toBe(10);
      expect(envelopeOf(captured.stdout())).toMatchObject({
        ok: false,
        error: { code: "locked" },
      });
      assertNoCanary(captured.stdout() + captured.stderr(), canary);
      await removeHeldLease(target);
    }

    // Text mode uses stderr for refusals while stdout remains empty.
    await writeHeldLease(file, canary);
    const textLocked = captureIo({ isTTY: true });
    expect(
      await runFolioCli(
        [
          "comment",
          file,
          "--input",
          JSON.stringify(argsFor("add_comment", revised)),
          "--expect-version",
          currentVersion,
          "--in-place",
        ],
        textLocked.io,
      ),
    ).toBe(10);
    expect(textLocked.stdout()).toBe("");
    expect(textLocked.stderr()).toContain("error:");
    assertNoCanary(textLocked.stdout() + textLocked.stderr(), canary);
    await removeHeldLease(file);

    // Independently exercise each registered MCP write tool through the MCP
    // protocol. This includes the compare tool's destination write variant.
    const [serverTransport, clientTransport] = InMemoryTransport.createLinkedPair();
    const server = createFolioMcpServer({ roots: [dir], author: "Canary Reviewer" });
    const client = new Client({ name: "lease-token-canary-test", version: "0.0.0" });
    await server.connect(serverTransport);
    await client.connect(clientTransport);
    try {
      const listed = await client.listTools();
      expect(listed.tools.map(({ name }) => name)).toContain("invoke_capability");
      for (const tool of writeTools) {
        const target = tool.type === "compare" ? path.join(dir, "mcp-redline.docx") : file;
        await writeHeldLease(target, canary);
        const result = await client.callTool({
          name: tool.name,
          arguments: {
            path: file,
            fileVersion: currentVersion,
            ...argsFor(tool.name, revised),
            ...(tool.type === "compare" && { destination: target }),
          },
        });
        const output = JSON.stringify(result);
        assertNoCanary(output, canary);
        const first = result.content.at(0);
        expect(
          first !== undefined && first.type === "text" ? JSON.parse(first.text) : null,
        ).toMatchObject({ error: { code: "locked" } });
        await removeHeldLease(target);
      }

      const readResult = await client.callTool({
        name: "read_document",
        arguments: { path: file },
      });
      assertNoCanary(JSON.stringify(readResult), canary);
      const statusResult = await client.callTool({
        name: "read_changes",
        arguments: { path: file },
      });
      assertNoCanary(JSON.stringify(statusResult), canary);
      for (const tool of readTools) {
        const result = await client.callTool({
          name: tool.name,
          arguments: { path: file, ...readArgsFor(tool.name, sectionHandle) },
        });
        assertNoCanary(JSON.stringify(result), canary);
      }
    } finally {
      await client.close();
      await server.close();
    }

    // A live editor that accepts flush requests leaves a pending request until
    // the wait deadline. Its lease token must stay out of the final CLI error.
    await writeHeldLease(file, canary, true);
    const flushWait = captureIo();
    expect(
      await runFolioCli(
        [
          "suggest",
          file,
          "--input",
          JSON.stringify(argsFor("suggest_changes", revised)),
          "--expect-version",
          fileVersionOf(new Uint8Array(await readFile(file))),
          "--in-place",
          "--flush-wait",
          "20",
        ],
        flushWait.io,
      ),
    ).toBe(10);
    expect(envelopeOf(flushWait.stdout())).toMatchObject({
      ok: false,
      error: { code: "locked", details: { flush: "timedOut" } },
    });
    assertNoCanary(flushWait.stdout() + flushWait.stderr(), canary);
    await removeHeldLease(file);

    await writeHeldLease(file, canary, true);
    const forced = await acquireLeaseForWrite({
      documentPath: file,
      txId: "force-canary-test",
      force: true,
      flushWaitMs: 20,
    });
    expect(forced.isOk()).toBe(true);
    if (forced.isOk()) {
      expect(forced.value.flush.type).toBe("timedOut");
      assertNoCanary(JSON.stringify(forced.value.flush), canary);
      await forced.value.lease.release();
    }

    const pendingId = randomUUID();
    const pendingPath = flushRequestPathFor(file, pendingId);
    await writeFile(
      pendingPath,
      JSON.stringify({
        id: pendingId,
        leaseToken: canary,
        owner: "canary-writer",
        pid: process.pid,
        host: hostname(),
        txId: "pending-canary-test",
        requestedAt: new Date().toISOString(),
        deadline: new Date(Date.now() + 60_000).toISOString(),
      }),
    );
    const pending = await acquireEditorLease({ documentPath: file, owner: "folio-test-editor" });
    expect(pending.isErr() && pending.error.code).toBe("locked");
    assertNoCanary(JSON.stringify(pending.isErr() ? pending.error : pending.value), canary);
    await rm(pendingPath, { force: true });
  }, 15_000);
});
