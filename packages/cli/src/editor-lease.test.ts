import { afterEach, beforeEach, describe, expect, test } from "bun:test";
import { readdir, readFile, writeFile } from "node:fs/promises";
import { hostname } from "node:os";
import path from "node:path";

import { FolioDocxReviewer } from "@stll/folio-core/server";

import { buildDocx, CONTRACT_PARAGRAPHS, makeTempDir, writeDocx } from "./__tests__/fixtures";
import { fileVersionOf } from "./document";
import {
  acquireEditorLease,
  acquireLeaseForWrite,
  flushRequestPathFor,
  keepEditorLeaseAlive,
  pendingFlushRequests,
  watchFlushRequests,
  type FlushWatcher,
} from "./editor-lease";
import { CLI_READ_BOUNDS, executeReadTool } from "./execute-read";
import { executeWriteTool, type WriteOptions } from "./execute-write";
import { lockPathFor, readLease, type AcquiredLease } from "./lock";
import { findFileTool } from "./registry";
import { saveDocumentBytes } from "./save";

let dir = "";
let cleanup: () => Promise<void> = () => Promise.resolve();
let file = "";
let watcher: FlushWatcher | undefined;
let editor: AcquiredLease | undefined;

/** A clause without a paraId, whose block id is derived from its text and position. */
const UNNUMBERED = { text: "Czech law governs this contract." };

beforeEach(async () => {
  ({ dir, cleanup } = await makeTempDir());
  file = await writeDocx(dir, "contract.docx", [...CONTRACT_PARAGRAPHS, UNNUMBERED]);
});

afterEach(async () => {
  watcher?.close();
  watcher = undefined;
  await editor?.release();
  editor = undefined;
  await cleanup();
});

const versionOf = async (filePath: string): Promise<string> =>
  fileVersionOf(new Uint8Array(await readFile(filePath)));

const REPLACE_FIFTY = {
  operations: [{ type: "replaceInBlock", blockId: "10000002", find: "$50", replace: "$500" }],
};

const write = async (
  toolName: string,
  args: Record<string, unknown>,
  fileVersion: string,
  overrides: Partial<WriteOptions> = {},
) => {
  const tool = findFileTool(toolName);
  if (!tool) throw new Error(`${toolName} is not registered`);
  return await executeWriteTool(
    tool,
    { path: file, fileVersion, args },
    {
      destination: { type: "inPlace" },
      author: "Agent",
      date: "2026-01-02T03:04:05Z",
      repack: "refuse",
      force: false,
      sourcePrecondition: "required",
      txId: undefined,
      journalPath: undefined,
      mode: "tracked-changes",
      // Generous, so a loaded CI machine never times out a flush that is coming.
      flushWaitMs: 30_000,
      ...overrides,
    },
  );
};

/** The user's unsaved edit: the third clause retyped. */
const unsavedEdit = (): Promise<Uint8Array> =>
  buildDocx([
    ...CONTRACT_PARAGRAPHS.slice(0, 2),
    { text: "Late payment accrues 2% interest.", paraId: "10000003" },
    ...CONTRACT_PARAGRAPHS.slice(3),
    UNNUMBERED,
  ]);

/** The block id `read` gives the paragraph without a paraId, checking it is synthetic. */
const syntheticIdOf = async (fileVersion: string): Promise<string> => {
  const tool = findFileTool("read_document");
  if (!tool) throw new Error("read_document is not registered");
  const data = (
    await executeReadTool(tool, { path: file, fileVersion, args: {} }, CLI_READ_BOUNDS)
  ).unwrap();
  const result: unknown = isObject(data) ? data["result"] : undefined;
  const blocks: unknown = isObject(result) ? result["blocks"] : undefined;
  const block = (Array.isArray(blocks) ? blocks : [])
    .filter(isObject)
    .find(({ text }) => text === UNNUMBERED.text);
  if (block?.["blockIdSource"] !== "synthetic" || typeof block["blockId"] !== "string") {
    throw new Error("no synthetic block id");
  }
  return block["blockId"];
};

const isObject = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

/**
 * An editor with unsaved edits that saves them and releases when asked;
 * `flushed` resolves to the version it saved.
 */
const startEditor = async (): Promise<{ flushed: Promise<string> }> => {
  const lease = (await acquireEditorLease({ documentPath: file, owner: "folio-vscode" })).unwrap();
  editor = lease;
  const loaded = await versionOf(file);
  const bytes = await unsavedEdit();
  let onRequest: () => void = () => undefined;
  const requested = new Promise<void>((resolve) => {
    onRequest = resolve;
  });
  watcher = watchFlushRequests({ documentPath: file, token: lease.holder.token, onRequest });
  const flush = async (): Promise<string> => {
    await requested;
    const saved = await saveDocumentBytes({
      path: file,
      bytes,
      expectedVersion: loaded,
      author: "User",
      owner: "folio-vscode",
      surface: "vscode",
      leaseToken: lease.holder.token,
    });
    await lease.release();
    return String(saved.unwrap()["fileVersion"]);
  };
  return { flushed: flush() };
};

const texts = async (filePath: string): Promise<string[]> => {
  const bytes = new Uint8Array(await readFile(filePath));
  const reviewer = await FolioDocxReviewer.fromBuffer(bytes.buffer);
  return reviewer.getContent().map(({ text }) => text);
};

describe("flush handshake", () => {
  test("an agent write waits for the editor to save, then applies on the saved version", async () => {
    const read = await versionOf(file);
    const { flushed } = await startEditor();

    const receipt = (await write("suggest_changes", REPLACE_FIFTY, read)).unwrap();
    const saved = await flushed;

    expect(receipt["status"]).toBe("committed");
    expect(receipt["fromVersion"]).toBe(saved);
    expect(receipt["rebased"]).toEqual({
      fromVersion: read,
      toVersion: saved,
      flushedBy: "folio-vscode",
    });
    const content = await texts(file);
    expect(content).toContain("Late payment accrues 2% interest.");
    expect(content.some((text) => text.includes("$500"))).toBe(true);
    const tools = (await readFile(path.join(dir, ".folio", "journal.jsonl"), "utf8"))
      .trim()
      .split("\n")
      .map((line) => JSON.parse(line).tool);
    expect(tools).toEqual(["editor_save", "suggest_changes"]);
    expect((await readdir(dir)).toSorted()).toEqual([".folio", "contract.docx"]);
  });

  test("a write targeting any text-derived block id is not carried over the flush", async () => {
    const read = await versionOf(file);
    const synthetic = await syntheticIdOf(read);
    const { flushed } = await startEditor();

    const result = await write(
      "suggest_changes",
      {
        operations: [
          ...REPLACE_FIFTY.operations,
          { type: "replaceInBlock", blockId: synthetic, find: "Czech", replace: "Slovak" },
        ],
      },
      read,
    );
    const saved = await flushed;

    expect(result.isErr() && result.error.code).toBe("stale_version");
    expect(await versionOf(file)).toBe(saved);
    const content = await texts(file);
    expect(content).toContain(UNNUMBERED.text);
    expect(content.some((text) => text.includes("$500"))).toBe(false);
  });

  test("a stale block precondition after the flush is the usual stale_target", async () => {
    const read = await versionOf(file);
    const { flushed } = await startEditor();

    const result = await write(
      "suggest_changes",
      {
        operations: [
          {
            type: "replaceInBlock",
            blockId: "10000003",
            find: "interest",
            replace: "a fee",
            precondition: { blockTextHash: "hstale" },
          },
        ],
      },
      read,
    );
    await flushed;

    expect(result.isErr() && result.error.code).toBe("stale_target");
    expect(await texts(file)).toContain("Late payment accrues 2% interest.");
  });

  test("tools whose targets are not block ids stay stale after a flush", async () => {
    const read = await versionOf(file);
    const { flushed } = await startEditor();

    const result = await write("resolve_changes", { action: "accept", all: true }, read);
    await flushed;

    expect(result.isErr() && result.error.code).toBe("stale_version");
  });

  test("a whole-package save from another owner flushes first, then is stale", async () => {
    const read = await versionOf(file);
    const { flushed } = await startEditor();

    const result = await saveDocumentBytes({
      path: file,
      bytes: await buildDocx(CONTRACT_PARAGRAPHS.slice(0, 2)),
      expectedVersion: read,
      author: "Other",
      owner: "folio-herdr",
      flushWaitMs: 30_000,
    });

    expect(await flushed).toBe(await versionOf(file));
    expect(result.isErr() && result.error.code).toBe("stale_version");
  });

  test("an unresponsive editor makes the write wait a bounded time, then refuse", async () => {
    editor = (await acquireEditorLease({ documentPath: file, owner: "folio-vscode" })).unwrap();
    const read = await versionOf(file);
    const started = Date.now();

    const result = await write("suggest_changes", REPLACE_FIFTY, read, { flushWaitMs: 300 });

    const waited = Date.now() - started;
    expect(result.isErr() && result.error.code).toBe("locked");
    expect(result.isErr() && JSON.stringify(result.error.details)).toContain('"timedOut"');
    expect(waited).toBeGreaterThanOrEqual(300);
    expect(waited).toBeLessThan(10_000);
    expect(await versionOf(file)).toBe(read);
    expect((await readdir(dir)).toSorted()).toEqual([".contract.docx.folio-lock", "contract.docx"]);
    expect((await editor.verify()).isOk()).toBe(true);
  });

  test("after the wait, --force takes the lease over from an unresponsive editor", async () => {
    editor = (await acquireEditorLease({ documentPath: file, owner: "folio-vscode" })).unwrap();

    const result = await write("suggest_changes", REPLACE_FIFTY, await versionOf(file), {
      flushWaitMs: 200,
      force: true,
    });

    expect(result.isOk()).toBe(true);
    const fenced = await editor.verify();
    expect(fenced.isErr() && fenced.error.code).toBe("locked");
  });

  test("a crashed editor's lease is replaced at once under the ordinary rules", async () => {
    const now = new Date();
    await writeFile(
      lockPathFor(file),
      JSON.stringify({
        owner: "folio-vscode",
        pid: 2 ** 22 + 12_345,
        host: hostname(),
        txId: "editor",
        token: "dead-editor",
        acquiredAt: now.toISOString(),
        expiresAt: new Date(now.getTime() + 30_000).toISOString(),
        acceptsFlush: true,
      }),
    );
    const started = Date.now();

    const acquired = (
      await acquireLeaseForWrite({ documentPath: file, txId: "tx", force: false })
    ).unwrap();

    expect(acquired.flush).toEqual({ type: "notAsked" });
    expect(Date.now() - started).toBeLessThan(1000);
    await acquired.lease.release();
  });
});

describe("editor lease", () => {
  test("is long-lived, advertises flushes, and renews", async () => {
    editor = (await acquireEditorLease({ documentPath: file, owner: "folio-vscode" })).unwrap();
    const lease = editor;
    const expires = Date.parse(lease.holder.expiresAt);

    const renewed = (await lease.renew(new Date(Date.now() + 5000))).unwrap();
    const state = await readLease(file);

    expect(lease.holder.acceptsFlush).toBe(true);
    expect(Date.parse(renewed.expiresAt)).toBeGreaterThan(expires);
    expect(state.type === "held" && state.holder).toMatchObject({
      owner: "folio-vscode",
      acceptsFlush: true,
      expiresAt: renewed.expiresAt,
    });
  });

  test("stops renewing and reports once the lease is taken over", async () => {
    editor = (await acquireEditorLease({ documentPath: file, owner: "folio-vscode" })).unwrap();
    const lost = new Promise<string>((resolve) => {
      if (!editor) throw new Error("no editor lease");
      keepEditorLeaseAlive(editor, { intervalMs: 20, onLost: (error) => resolve(error.code) });
    });

    (
      await acquireLeaseForWrite({ documentPath: file, txId: "t", force: true, flushWaitMs: 0 })
    ).unwrap();

    expect(await lost).toBe("locked");
  });

  test("is refused while a writer's flush request is pending", async () => {
    const deadline = new Date(Date.now() + 60_000).toISOString();
    await writeFile(
      flushRequestPathFor(file, "r1"),
      JSON.stringify({
        id: "r1",
        leaseToken: "old",
        owner: "folio-cli",
        pid: process.pid,
        host: hostname(),
        txId: "tx",
        requestedAt: new Date().toISOString(),
        deadline,
      }),
    );

    const refused = await acquireEditorLease({ documentPath: file, owner: "folio-vscode" });

    expect(refused.isErr() && refused.error.code).toBe("locked");
    expect((await pendingFlushRequests(file)).map(({ id }) => id)).toEqual(["r1"]);
    expect(await pendingFlushRequests(file, new Date(Date.parse(deadline) + 1))).toEqual([]);
    expect(await readdir(dir)).toEqual(["contract.docx"]);
  });
});
