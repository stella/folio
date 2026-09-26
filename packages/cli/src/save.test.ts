import { afterEach, beforeEach, describe, expect, test } from "bun:test";
import { readdir, readFile, rm, writeFile } from "node:fs/promises";
import path from "node:path";

import { FolioDocxReviewer } from "@stll/folio-core/server";

import {
  buildDocx,
  CONTRACT_PARAGRAPHS,
  makeTempDir,
  writeDocx,
  type FixtureParagraph,
} from "./__tests__/fixtures";
import { captureIo, dataOf, envelopeOf } from "./__tests__/io";
import { runFolioCli } from "./cli";
import { fileVersionOf, MAX_DOCUMENT_BYTES } from "./document";
import { acquireEditorLease } from "./editor-lease";
import { latestCommitFor, stagePathFor } from "./journal";
import { acquireLease, readLease } from "./lock";
import { listMcpTools } from "./mcp";
import { saveDocumentBytes, type SaveDocumentOptions } from "./save";

let dir = "";
let cleanup: () => Promise<void> = () => Promise.resolve();
let file = "";

beforeEach(async () => {
  ({ dir, cleanup } = await makeTempDir());
  file = await writeDocx(dir, "contract.docx", CONTRACT_PARAGRAPHS);
});

afterEach(async () => {
  await cleanup();
});

const DATE = "2026-01-02T03:04:05Z";

const versionOf = async (filePath: string): Promise<string> =>
  fileVersionOf(new Uint8Array(await readFile(filePath)));

/** The contract with its third paragraph retyped, and optionally a new last one. */
const editedBytes = (extra: readonly FixtureParagraph[] = []): Promise<Uint8Array> =>
  buildDocx([
    ...CONTRACT_PARAGRAPHS.slice(0, 2),
    { text: "Late payment accrues 2% interest.", paraId: "10000003" },
    ...CONTRACT_PARAGRAPHS.slice(3),
    ...extra,
  ]);

const save = async (overrides: Partial<SaveDocumentOptions> = {}) =>
  await saveDocumentBytes({
    path: file,
    bytes: await editedBytes(),
    expectedVersion: await versionOf(file),
    author: "Editor",
    surface: "vscode",
    saveStrategy: "full-repack",
    date: DATE,
    ...overrides,
  });

const texts = async (filePath: string): Promise<string[]> => {
  const bytes = new Uint8Array(await readFile(filePath));
  const reviewer = await FolioDocxReviewer.fromBuffer(bytes.buffer);
  return reviewer.getContent().map(({ text }) => text);
};

describe("saveDocumentBytes", () => {
  test("commits the editor's package with a backup and an editor_save journal line", async () => {
    const before = await versionOf(file);

    const receipt = (await save({ txId: "save-1" })).unwrap();

    expect(receipt["status"]).toBe("committed");
    expect(receipt["tool"]).toBe("editor_save");
    expect(receipt["fromVersion"]).toBe(before);
    expect(receipt["fileVersion"]).toBe(await versionOf(file));
    expect(receipt["changedParts"]).toEqual([{ part: "word/document.xml", change: "modified" }]);
    expect(await versionOf(String(receipt["backup"]))).toBe(before);
    expect(await texts(file)).toContain("Late payment accrues 2% interest.");
    const line = JSON.parse(
      (await readFile(path.join(dir, ".folio", "journal.jsonl"), "utf8")).trim(),
    );
    expect(line).toMatchObject({
      type: "commit",
      txId: "save-1",
      tool: "editor_save",
      author: "Editor",
      fromVersion: before,
      toVersion: receipt["fileVersion"],
      ops: { surface: "vscode", saveStrategy: "full-repack", owner: "folio-cli" },
    });
    expect(await latestCommitFor(file)).toEqual({
      txId: "save-1",
      tool: "editor_save",
      author: "Editor",
      time: DATE,
      fromVersion: before,
      toVersion: String(receipt["fileVersion"]),
    });
    expect((await readdir(dir)).toSorted()).toEqual([".folio", "contract.docx"]);
    expect((await readLease(file)).type).toBe("free");
  });

  test("writes a structural change (a full rewrite) and replays its txId", async () => {
    const bytes = await editedBytes([{ text: "A new clause.", paraId: "10000005" }]);

    const first = (await save({ bytes, txId: "save-2" })).unwrap();
    const replayed = (
      await save({ bytes, txId: "save-2", expectedVersion: String(first["fromVersion"]) })
    ).unwrap();

    expect(await texts(file)).toContain("A new clause.");
    expect(replayed).toEqual({ ...first, status: "replayed" });
  });

  test("an unchanged package writes nothing", async () => {
    const bytes = new Uint8Array(await readFile(file));

    const receipt = (await save({ bytes })).unwrap();

    expect(receipt["status"]).toBe("unchanged");
    expect(await readdir(dir)).toEqual(["contract.docx"]);
  });

  test("saves to another file with -o, never touching the source", async () => {
    const out = path.join(dir, "copy.docx");
    const before = await versionOf(file);

    const receipt = (
      await save({
        destination: { type: "file", path: out, overwrite: false, expectedVersion: undefined },
      })
    ).unwrap();

    expect(receipt["source"]).toEqual({ path: file, fileVersion: before });
    expect(await versionOf(file)).toBe(before);
    expect(await texts(out)).toContain("Late payment accrues 2% interest.");
  });
});

describe("refusals leave the file untouched", () => {
  test.each([
    ["stale version", () => save({ expectedVersion: "0".repeat(64) }), "stale_version"],
    [
      "bytes that are not a package",
      () => save({ bytes: new TextEncoder().encode("not a zip") }),
      "invalid_document",
    ],
    [
      "bytes over 64 MiB",
      () => save({ bytes: new Uint8Array(MAX_DOCUMENT_BYTES + 1) }),
      "too_large",
    ],
    ["an unsafe owner", () => save({ owner: "../x" }), "invalid_input"],
  ] as const)("%s", async (_name, run, code) => {
    const before = await versionOf(file);

    const result = await run();

    expect(result.isErr() && result.error.code).toBe(code);
    expect(await versionOf(file)).toBe(before);
    expect(await readdir(dir)).not.toContain(".folio");
  });

  test("a lease another process holds refuses", async () => {
    const other = (
      await acquireLease({ documentPath: file, txId: "agent", force: false })
    ).unwrap();
    const before = await versionOf(file);

    const result = await save();

    expect(result.isErr() && result.error.code).toBe("locked");
    expect(await versionOf(file)).toBe(before);
    await other.release();
  });
});

describe("under the editor's lease", () => {
  test("--lease-token saves under a held lease and leaves it held", async () => {
    const lease = (
      await acquireEditorLease({ documentPath: file, owner: "folio-vscode" })
    ).unwrap();

    const receipt = (await save({ leaseToken: lease.holder.token })).unwrap();
    const wrongToken = await save({
      leaseToken: "not-the-token",
      expectedVersion: String(receipt["fileVersion"]),
    });

    expect(receipt["status"]).toBe("committed");
    expect((await lease.verify()).isOk()).toBe(true);
    expect(wrongToken.isErr() && wrongToken.error.code).toBe("locked");
    await lease.release();
  });
});

describe("crash recovery", () => {
  test("rolls a journaled stage forward, then checks the save against the result", async () => {
    const original = await readFile(file);
    const first = (await save({ txId: "tx-1" })).unwrap();
    const committed = await readFile(file);
    // Simulate a crash between the journal line and the rename.
    await rm(file);
    await writeFile(file, original);
    await writeFile(stagePathFor(file, "tx-1"), committed);
    const next = await editedBytes([{ text: "Signed in duplicate.", paraId: "10000005" }]);

    const stale = await save({ bytes: next, expectedVersion: fileVersionOf(original) });
    const saved = (
      await save({ bytes: next, expectedVersion: String(first["fileVersion"]) })
    ).unwrap();

    expect(stale.isErr() && stale.error.code).toBe("stale_version");
    expect(stale.isErr() && JSON.stringify(stale.error.details)).toContain('"rolledForward"');
    expect(saved["fromVersion"]).toBe(fileVersionOf(new Uint8Array(committed)));
    expect(await texts(file)).toContain("Signed in duplicate.");
  });
});

describe("folio save", () => {
  test("commits --from with the envelope and exit codes", async () => {
    const from = path.join(dir, "edited.docx");
    await writeFile(from, await editedBytes());
    const before = await versionOf(file);
    const ok = captureIo();
    const stale = captureIo();

    const argv = ["save", file, "--from", from, "--surface", "vscode", "--owner", "folio-vscode"];
    const okExit = await runFolioCli(
      [...argv, "--expect-version", before, "--output", "json"],
      ok.io,
    );
    const staleExit = await runFolioCli([...argv, "--expect-version", before], stale.io);

    expect(okExit).toBe(0);
    expect(dataOf(ok.stdout())).toMatchObject({
      status: "committed",
      tool: "editor_save",
      owner: "folio-vscode",
      surface: "vscode",
      fromVersion: before,
    });
    expect(staleExit).toBe(10);
    expect(JSON.stringify(envelopeOf(stale.stdout())["error"])).toContain('"stale_version"');
  });

  test("refuses a call without --from or --expect-version", async () => {
    for (const argv of [
      ["save", file, "--expect-version", "0".repeat(64)],
      ["save", file, "--from", file],
    ]) {
      const captured = captureIo();
      expect(await runFolioCli(argv, captured.io)).toBe(2);
      expect(JSON.stringify(envelopeOf(captured.stdout())["error"])).toContain('"usage_error"');
    }
  });

  test("is listed in help but never offered over MCP", async () => {
    const captured = captureIo();
    await runFolioCli(["--help"], captured.io);

    expect(captured.stdout()).toContain("save");
    const names = listMcpTools().map(({ name }) => name);
    expect(names).not.toContain("save");
    expect(names).not.toContain("editor_save");
  });
});
