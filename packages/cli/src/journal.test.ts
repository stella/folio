import { afterEach, beforeEach, describe, expect, test } from "bun:test";
import { appendFile, mkdir, readdir, readFile, symlink, writeFile } from "node:fs/promises";
import path from "node:path";

import { buildDocx, CONTRACT_PARAGRAPHS, makeTempDir } from "./__tests__/fixtures";
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
let original = new Uint8Array();
let staged = new Uint8Array();

beforeEach(async () => {
  ({ dir, cleanup } = await makeTempDir());
  documentPath = path.join(dir, "contract.docx");
  journalPath = journalPathFor(documentPath, undefined);
  original = await buildDocx(CONTRACT_PARAGRAPHS);
  staged = await buildDocx([{ text: "Staged.", paraId: "20000001" }]);
  await writeFile(documentPath, original);
});

afterEach(async () => {
  await cleanup();
});

type CommitLineOptions = {
  txId: string;
  destinationVersionBefore: string | null;
  path?: string;
  toVersion?: string;
};

const commitLine = ({
  txId,
  destinationVersionBefore,
  path: committedPath = documentPath,
  toVersion = fileVersionOf(staged),
}: CommitLineOptions): JournalCommit => ({
  type: "commit",
  txId,
  requestHash: "request",
  tool: "suggest_changes",
  path: committedPath,
  fromVersion: fileVersionOf(original),
  toVersion,
  destinationVersionBefore,
  stage: path.basename(stagePathFor(documentPath, txId)),
  author: "Reviewer",
  time: "2026-01-02T03:04:05Z",
  ops: {},
  receipts: [],
  receipt: { txId, status: "committed" },
});

const now = "2026-01-02T03:05:00Z";
const recover = () => recoverStages({ documentPath, journalPath, now });
const listing = async () => (await readdir(dir)).toSorted();

describe("recoverStages", () => {
  test("rolls a journaled stage forward while the destination is unchanged", async () => {
    await writeFile(stagePathFor(documentPath, "tx-1"), staged);
    const line = commitLine({ txId: "tx-1", destinationVersionBefore: fileVersionOf(original) });
    (await appendJournal(journalPath, line)).unwrap();

    const actions = (await recover()).unwrap();

    expect(actions).toEqual([{ txId: "tx-1", action: "rolledForward" }]);
    expect(fileVersionOf(new Uint8Array(await readFile(documentPath)))).toBe(fileVersionOf(staged));
    expect(await listing()).toEqual([".folio", "contract.docx"]);
    expect((await findCommit(journalPath, "tx-1")).unwrap()?.toVersion).toBe(fileVersionOf(staged));
  });

  test("discards stages that are not journaled, name another document, or were outrun", async () => {
    const before = fileVersionOf(original);
    await writeFile(stagePathFor(documentPath, "never-journaled"), staged);
    await writeFile(stagePathFor(documentPath, "outran"), staged);
    await writeFile(stagePathFor(documentPath, "elsewhere"), staged);
    (
      await appendJournal(
        journalPath,
        commitLine({ txId: "outran", destinationVersionBefore: "x" }),
      )
    ).unwrap();
    const foreign = commitLine({
      txId: "elsewhere",
      destinationVersionBefore: before,
      path: path.join(dir, "other.docx"),
    });
    (await appendJournal(journalPath, foreign)).unwrap();

    const actions = (await recover()).unwrap();

    expect(actions.toSorted((left, right) => left.txId.localeCompare(right.txId))).toEqual([
      { txId: "elsewhere", action: "discarded", reason: "otherDocument" },
      { txId: "outran", action: "discarded", reason: "destinationMoved" },
    ]);
    expect(fileVersionOf(new Uint8Array(await readFile(documentPath)))).toBe(before);
    expect(await listing()).toEqual([".folio", "contract.docx"]);
    expect((await findCommit(journalPath, "outran")).unwrap()).toBeUndefined();
  });

  test("never follows a symlinked stage and leaves its target alone", async () => {
    const target = path.join(dir, "planted.docx");
    await writeFile(target, staged);
    await symlink(target, stagePathFor(documentPath, "tx-link"));
    const line = commitLine({ txId: "tx-link", destinationVersionBefore: fileVersionOf(original) });
    (await appendJournal(journalPath, line)).unwrap();

    const actions = (await recover()).unwrap();

    expect(actions).toEqual([{ txId: "tx-link", action: "discarded", reason: "notRegularFile" }]);
    expect(fileVersionOf(new Uint8Array(await readFile(documentPath)))).toBe(
      fileVersionOf(original),
    );
    expect(await listing()).toEqual([".folio", "contract.docx", "planted.docx"]);
  });

  test("discards a stage whose bytes are not a package", async () => {
    const garbage = new TextEncoder().encode("not a package");
    await writeFile(stagePathFor(documentPath, "tx-bad"), garbage);
    const line = commitLine({
      txId: "tx-bad",
      destinationVersionBefore: fileVersionOf(original),
      toVersion: fileVersionOf(garbage),
    });
    (await appendJournal(journalPath, line)).unwrap();

    expect((await recover()).unwrap()).toEqual([
      { txId: "tx-bad", action: "discarded", reason: "invalidPackage" },
    ]);
  });
});

describe("journal boundary", () => {
  test("skips a torn final line and other transactions", async () => {
    (
      await appendJournal(journalPath, commitLine({ txId: "tx-1", destinationVersionBefore: null }))
    ).unwrap();
    await appendFile(journalPath, '{"type":"commit","txId":"tx-2","requ');

    expect((await findCommit(journalPath, "tx-1")).unwrap()?.receipt).toEqual({
      txId: "tx-1",
      status: "committed",
    });
    expect((await findCommit(journalPath, "tx-2")).unwrap()).toBeUndefined();
    expect((await findCommit(path.join(dir, "missing.jsonl"), "tx-1")).unwrap()).toBeUndefined();
  });

  test("refuses to write or read through a symlinked .folio or journal", async () => {
    const { dir: elsewhere, cleanup: cleanupElsewhere } = await makeTempDir();
    try {
      await symlink(elsewhere, path.join(dir, ".folio"));
      const line = commitLine({ txId: "tx-1", destinationVersionBefore: null });

      const appended = await appendJournal(journalPath, line);
      const found = await findCommit(journalPath, "tx-1");

      expect(appended.isErr() && appended.error.code).toBe("unsafe_path");
      expect(found.isErr() && found.error.code).toBe("unsafe_path");
      expect(await readdir(elsewhere)).toEqual([]);

      const separate = path.join(dir, "separate");
      await mkdir(separate);
      await symlink(path.join(elsewhere, "journal.jsonl"), path.join(separate, "journal.jsonl"));
      const redirected = await appendJournal(path.join(separate, "journal.jsonl"), line);
      expect(redirected.isErr() && redirected.error.code).toBe("unsafe_path");
      expect(await readdir(elsewhere)).toEqual([]);
    } finally {
      await cleanupElsewhere();
    }
  });
});
