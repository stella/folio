import { afterEach, beforeEach, describe, expect, test } from "bun:test";
import { readFile, writeFile } from "node:fs/promises";
import { hostname } from "node:os";
import path from "node:path";

import { makeTempDir } from "./__tests__/fixtures";
import { acquireLease, lockPathFor, readLease, type LockHolder } from "./lock";

let dir = "";
let cleanup: () => Promise<void> = () => Promise.resolve();
let documentPath = "";

beforeEach(async () => {
  ({ dir, cleanup } = await makeTempDir());
  documentPath = path.join(dir, "contract.docx");
});

afterEach(async () => {
  await cleanup();
});

const writeHolder = async (holder: Partial<LockHolder>): Promise<void> => {
  const now = new Date();
  await writeFile(
    lockPathFor(documentPath),
    JSON.stringify({
      owner: "editor",
      pid: process.pid,
      host: hostname(),
      txId: "other",
      acquiredAt: now.toISOString(),
      expiresAt: new Date(now.getTime() + 60_000).toISOString(),
      ...holder,
    }),
  );
};

describe("acquireLease", () => {
  test("holds the lease beside the document and releases only its own", async () => {
    const lease = (await acquireLease({ documentPath, txId: "tx-1", force: false })).unwrap();

    const second = await acquireLease({ documentPath, txId: "tx-2", force: false });
    expect(second.isErr() && second.error.code).toBe("locked");
    expect(path.basename(lockPathFor(documentPath))).toBe(".contract.docx.folio-lock");

    await lease.release();
    expect((await readLease(documentPath)).type).toBe("free");
  });

  test("refuses a live holder unless forced, and leaves its lease on release", async () => {
    await writeHolder({});

    const refused = await acquireLease({ documentPath, txId: "tx-1", force: false });
    const forced = (await acquireLease({ documentPath, txId: "tx-1", force: true })).unwrap();

    expect(refused.isErr() && refused.error.code).toBe("locked");
    expect(forced.holder.txId).toBe("tx-1");
    await writeHolder({});
    await forced.release();
    expect((await readLease(documentPath)).type).toBe("held");
  });

  test("replaces an expired lease and one whose process has exited", async () => {
    for (const stale of [
      { expiresAt: new Date(Date.now() - 1000).toISOString() },
      { pid: 2 ** 22 + 12_345 },
    ]) {
      await writeHolder(stale);

      const lease = await acquireLease({ documentPath, txId: "tx-new", force: false });

      expect(lease.isOk()).toBe(true);
      expect(await readFile(lockPathFor(documentPath), "utf8")).toContain('"txId":"tx-new"');
      await lease.unwrap().release();
    }
  });
});
