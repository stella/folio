import { afterEach, beforeEach, describe, expect, test } from "bun:test";
import { readFile, utimes, writeFile } from "node:fs/promises";
import { hostname } from "node:os";
import path from "node:path";

import { makeTempDir } from "./__tests__/fixtures";
import {
  acquireLease,
  lockPathFor,
  readLease,
  TRANSACTION_LEASE_MS,
  type LockHolder,
} from "./lock";

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
      token: "other-token",
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
    expect((await lease.verify()).isOk()).toBe(true);

    await lease.release();
    expect((await readLease(documentPath)).type).toBe("free");
  });

  test("a forced takeover fences out the previous holder", async () => {
    const first = (await acquireLease({ documentPath, txId: "tx-1", force: false })).unwrap();

    const refused = await acquireLease({ documentPath, txId: "tx-2", force: false });
    const forced = (await acquireLease({ documentPath, txId: "tx-2", force: true })).unwrap();

    expect(refused.isErr() && refused.error.code).toBe("locked");
    const fenced = await first.verify();
    expect(fenced.isErr() && fenced.error.code).toBe("locked");
    await first.release();
    expect((await forced.verify()).isOk()).toBe(true);
  });

  test("does not take over a fresh lock it cannot parse", async () => {
    await writeFile(lockPathFor(documentPath), '{"owner":"edi');

    const refused = await acquireLease({ documentPath, txId: "tx-1", force: false });

    expect(refused.isErr() && refused.error.code).toBe("locked");
    expect(await readFile(lockPathFor(documentPath), "utf8")).toBe('{"owner":"edi');
  });

  test("replaces an expired lease, one whose process exited, and an old unparsable lock", async () => {
    const old = new Date(Date.now() - TRANSACTION_LEASE_MS - 60_000);
    for (const plant of [
      () => writeHolder({ expiresAt: new Date(Date.now() - 1000).toISOString() }),
      () => writeHolder({ pid: 2 ** 22 + 12_345 }),
      async () => {
        await writeFile(lockPathFor(documentPath), "garbage");
        await utimes(lockPathFor(documentPath), old, old);
      },
    ]) {
      await plant();

      const lease = await acquireLease({ documentPath, txId: "tx-new", force: false });

      expect(lease.isOk()).toBe(true);
      expect(await readFile(lockPathFor(documentPath), "utf8")).toContain('"txId":"tx-new"');
      await lease.unwrap().release();
    }
  });
});
