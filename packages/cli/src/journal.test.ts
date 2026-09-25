import { afterEach, beforeEach, describe, expect, test } from "bun:test";
import { appendFile, readdir, readFile, writeFile } from "node:fs/promises";
import path from "node:path";

import { makeTempDir } from "./__tests__/fixtures";
import { fileVersionOf } from "./document";
import {
  appendJournal,
  findCommit,
  journalPathFor,
  recoverStages,
  stagePathFor,
  type JournalCommit,
} from "./journal";

let dir = "";
let cleanup: () => Promise<void> = () => Promise.resolve();
let documentPath = "";
let journalPath = "";

const ORIGINAL = new TextEncoder().encode("original bytes");
const STAGED = new TextEncoder().encode("staged bytes");

beforeEach(async () => {
  ({ dir, cleanup } = await makeTempDir());
  documentPath = path.join(dir, "contract.docx");
  journalPath = journalPathFor(documentPath, undefined);
  await writeFile(documentPath, ORIGINAL);
});

afterEach(async () => {
  await cleanup();
});

const commitLine = (txId: string, destinationVersionBefore: string | null): JournalCommit => ({
  type: "commit",
  txId,
  requestHash: "request",
  tool: "suggest_changes",
  path: documentPath,
  fromVersion: fileVersionOf(ORIGINAL),
  toVersion: fileVersionOf(STAGED),
  destinationVersionBefore,
  stage: path.basename(stagePathFor(documentPath, txId)),
  author: "Reviewer",
  time: "2026-01-02T03:04:05Z",
  ops: {},
  receipts: [],
  receipt: { txId, status: "committed" },
});

const now = "2026-01-02T03:05:00Z";

describe("recoverStages", () => {
  test("rolls a journaled stage forward while the destination is unchanged", async () => {
    await writeFile(stagePathFor(documentPath, "tx-1"), STAGED);
    (await appendJournal(journalPath, commitLine("tx-1", fileVersionOf(ORIGINAL)))).unwrap();

    const actions = (await recoverStages({ documentPath, journalPath, now })).unwrap();

    expect(actions).toEqual([{ txId: "tx-1", action: "rolledForward" }]);
    expect(await readFile(documentPath, "utf8")).toBe("staged bytes");
    expect((await readdir(dir)).toSorted()).toEqual([".folio", "contract.docx"]);
    expect((await findCommit(journalPath, "tx-1"))?.toVersion).toBe(fileVersionOf(STAGED));
  });

  test("discards a stage the journal never committed or the destination outran", async () => {
    await writeFile(stagePathFor(documentPath, "never-journaled"), STAGED);
    await writeFile(stagePathFor(documentPath, "outran"), STAGED);
    (await appendJournal(journalPath, commitLine("outran", "another version"))).unwrap();

    const actions = (await recoverStages({ documentPath, journalPath, now })).unwrap();

    expect(actions).toEqual([{ txId: "outran", action: "discarded" }]);
    expect(await readFile(documentPath, "utf8")).toBe("original bytes");
    expect((await readdir(dir)).toSorted()).toEqual([".folio", "contract.docx"]);
    expect(await findCommit(journalPath, "outran")).toBeUndefined();
  });
});

describe("findCommit", () => {
  test("skips a torn final line and other transactions", async () => {
    (await appendJournal(journalPath, commitLine("tx-1", null))).unwrap();
    await appendFile(journalPath, '{"type":"commit","txId":"tx-2","requ');

    expect((await findCommit(journalPath, "tx-1"))?.receipt).toEqual({
      txId: "tx-1",
      status: "committed",
    });
    expect(await findCommit(journalPath, "tx-2")).toBeUndefined();
    expect(await findCommit(path.join(dir, "missing.jsonl"), "tx-1")).toBeUndefined();
  });
});
