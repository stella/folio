import { afterEach, beforeEach, describe, expect, test } from "bun:test";
import { readdir, readFile, writeFile } from "node:fs/promises";
import path from "node:path";

import { FolioDocxReviewer } from "@stll/folio-core/server";

import { CONTRACT_PARAGRAPHS, makeTempDir, writeDocx } from "./__tests__/fixtures";
import { fileVersionOf } from "./document";
import { executeWriteTool, type WriteOptions } from "./execute-write";
import { acquireLease } from "./lock";
import { findFileTool } from "./registry";
import { stagePathFor } from "./journal";

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

const options = (overrides: Partial<WriteOptions> = {}): WriteOptions => ({
  destination: { type: "inPlace" },
  author: "Reviewer",
  date: DATE,
  repack: "refuse",
  force: false,
  txId: undefined,
  journalPath: undefined,
  mode: "tracked-changes",
  ...overrides,
});

const write = async (
  toolName: string,
  args: Record<string, unknown>,
  overrides: Partial<WriteOptions> = {},
  fileVersion?: string,
) => {
  const tool = findFileTool(toolName);
  if (!tool) throw new Error(`${toolName} is not registered`);
  return await executeWriteTool(tool, { path: file, fileVersion, args }, options(overrides));
};

const REPLACE_FIFTY = {
  operations: [{ type: "replaceInBlock", blockId: "10000002", find: "$50", replace: "$500" }],
};

const versionOf = async (filePath: string): Promise<string> =>
  fileVersionOf(new Uint8Array(await readFile(filePath)));

const reopen = async (filePath: string): Promise<FolioDocxReviewer> => {
  const bytes = new Uint8Array(await readFile(filePath));
  return await FolioDocxReviewer.fromBuffer(bytes.buffer);
};

describe("suggest_changes in place", () => {
  test("commits a selective save with a backup, a journal line, and stamped revisions", async () => {
    const before = await versionOf(file);

    const receipt = (await write("suggest_changes", REPLACE_FIFTY, {}, before)).unwrap();

    expect(receipt["status"]).toBe("committed");
    expect(receipt["fromVersion"]).toBe(before);
    expect(receipt["fileVersion"]).toBe(await versionOf(file));
    expect(receipt["saveStrategy"]).toBe("selective");
    expect(receipt["changedParts"]).toEqual([{ part: "word/document.xml", change: "modified" }]);
    expect(await versionOf(String(receipt["backup"]))).toBe(before);
    const changes = (await reopen(file)).getChanges();
    expect(changes.map(({ author, date }) => [author, date])).toEqual([
      ["Reviewer", DATE],
      ["Reviewer", DATE],
    ]);
    const journal = await readFile(path.join(dir, ".folio", "journal.jsonl"), "utf8");
    expect(journal).toContain(`"toVersion":"${String(receipt["fileVersion"])}"`);
    expect((await readdir(dir)).toSorted()).toEqual([".folio", "contract.docx"]);
  });

  test("replays a committed txId and refuses its reuse for another request", async () => {
    const first = (await write("suggest_changes", REPLACE_FIFTY, { txId: "tx-1" })).unwrap();
    const after = await versionOf(file);

    const replayed = (await write("suggest_changes", REPLACE_FIFTY, { txId: "tx-1" })).unwrap();
    const conflict = await write(
      "suggest_changes",
      { operations: [{ type: "deleteBlock", blockId: "10000004" }] },
      { txId: "tx-1" },
    );

    expect(replayed).toEqual({ ...first, status: "replayed" });
    expect(await versionOf(file)).toBe(after);
    expect(conflict.isErr() && conflict.error.code).toBe("transaction_conflict");
  });
});

describe("refusals leave the file untouched", () => {
  const cases: [string, () => ReturnType<typeof write>, string][] = [
    [
      "stale version",
      () => write("suggest_changes", REPLACE_FIFTY, {}, "0".repeat(64)),
      "stale_version",
    ],
    [
      "stale block",
      () =>
        write("suggest_changes", {
          operations: [
            {
              type: "replaceInBlock",
              blockId: "10000002",
              find: "$50",
              replace: "$5",
              precondition: { blockTextHash: "hstale" },
            },
          ],
        }),
      "stale_target",
    ],
    [
      "ambiguous find",
      () =>
        write("suggest_changes", {
          operations: [{ type: "replaceInBlock", blockId: "10000002", find: "e", replace: "E" }],
        }),
      "ambiguous_target",
    ],
    [
      "structural edit without --allow-repack",
      () =>
        write("suggest_changes", {
          operations: [{ type: "insertAfterBlock", blockId: "10000004", text: "New clause." }],
        }),
      "repack_required",
    ],
    ["unknown comment", () => write("reply_comment", { commentId: "404", text: "?" }), "not_found"],
    [
      "unknown change",
      () => write("resolve_changes", { action: "accept", ids: ["9"] }),
      "not_found",
    ],
    [
      "nothing to accept",
      () => write("resolve_changes", { action: "accept", all: true }),
      "operation_rejected",
    ],
    [
      "ids and all together",
      () => write("resolve_changes", { action: "accept", all: true, ids: ["1"] }),
      "invalid_input",
    ],
    [
      "unsafe txId",
      () => write("suggest_changes", REPLACE_FIFTY, { txId: "../escape" }),
      "invalid_input",
    ],
  ];

  test.each(cases)("%s", async (_name, run, code) => {
    const before = await versionOf(file);

    const result = await run();

    expect(result.isErr() && result.error.code).toBe(code);
    expect(await versionOf(file)).toBe(before);
  });

  test("a held lease refuses unless forced", async () => {
    const other = (
      await acquireLease({ documentPath: file, txId: "editor", force: false })
    ).unwrap();

    const refused = await write("suggest_changes", REPLACE_FIFTY);
    const forced = await write("suggest_changes", REPLACE_FIFTY, { force: true });

    expect(refused.isErr() && refused.error.code).toBe("locked");
    expect(forced.isOk()).toBe(true);
    await other.release();
  });
});

describe("destinations", () => {
  test("-o refuses an existing file unless overwriting, and never touches the source", async () => {
    const out = path.join(dir, "out.docx");
    await writeFile(out, "occupied");
    const source = await versionOf(file);

    const refused = await write("suggest_changes", REPLACE_FIFTY, {
      destination: { type: "file", path: out, overwrite: false },
    });
    const replaced = await write("suggest_changes", REPLACE_FIFTY, {
      destination: { type: "file", path: out, overwrite: true },
    });

    expect(refused.isErr() && refused.error.code).toBe("destination_exists");
    expect(replaced.unwrap()["source"]).toEqual({ path: file, fileVersion: source });
    expect(await versionOf(file)).toBe(source);
    expect((await reopen(out)).getChanges().length).toBe(2);
  });

  test("a full repack is reported when allowed", async () => {
    const receipt = (
      await write(
        "suggest_changes",
        { operations: [{ type: "insertAfterBlock", blockId: "10000004", text: "New clause." }] },
        { repack: "allow" },
      )
    ).unwrap();

    expect(receipt["saveStrategy"]).toBe("full-repack");
    expect(receipt["repackReason"]).toBe("structuralChange");
  });
});

describe("resolve_changes", () => {
  test("accepts by id and then all remaining", async () => {
    (await write("suggest_changes", REPLACE_FIFTY)).unwrap();
    const [first] = (await reopen(file)).getChanges();
    if (!first) throw new Error("no change to accept");

    const byId = (
      await write("resolve_changes", { action: "accept", ids: [String(first.id)] })
    ).unwrap();
    const rest = (await write("resolve_changes", { action: "reject", all: true })).unwrap();

    expect(byId["result"]).toEqual({
      action: "accept",
      resolved: [String(first.id)],
      remaining: 1,
    });
    expect(rest["result"]).toEqual({ action: "reject", resolved: 1, remaining: 0 });
    expect((await reopen(file)).getContent().map(({ text }) => text)).toContain(
      "The buyer pays  on signing.",
    );
  });
});

describe("crash recovery", () => {
  test("a later write rolls a journaled stage forward before applying", async () => {
    (await write("suggest_changes", REPLACE_FIFTY, { txId: "tx-1" })).unwrap();
    const committed = await readFile(file);
    const backups = await readdir(path.join(dir, ".folio", "backups"));
    const original = await readFile(path.join(dir, ".folio", "backups", backups[0] ?? ""));
    // Simulate a crash between the journal line and the rename.
    await writeFile(file, original);
    await writeFile(stagePathFor(file, "tx-1"), committed);

    const next = (
      await write("add_comment", { blockId: "10000003", text: "Confirm the rate." })
    ).unwrap();

    expect(next["recovered"]).toEqual([{ txId: "tx-1", action: "rolledForward" }]);
    expect(next["fromVersion"]).toBe(fileVersionOf(new Uint8Array(committed)));
    const reviewer = await reopen(file);
    expect(reviewer.getChanges().length).toBe(2);
    expect(reviewer.getComments().map(({ text }) => text)).toEqual(["Confirm the rate."]);
  });
});
