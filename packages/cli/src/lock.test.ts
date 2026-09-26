import { afterEach, beforeEach, describe, expect, test } from "bun:test";
import { readdir, readFile, rm, utimes, writeFile } from "node:fs/promises";
import { hostname } from "node:os";
import path from "node:path";
import { setTimeout as sleep } from "node:timers/promises";

import { makeTempDir } from "./__tests__/fixtures";
import {
  acquireLease,
  lockPathFor,
  lockSwapPathFor,
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

const tokenOnDisk = async (): Promise<string | null> => {
  const state = await readLease(documentPath);
  return state.type === "held" ? state.holder.token : null;
};

describe("renewal against a takeover", () => {
  test("a renewal queued behind a takeover refuses instead of overwriting it", async () => {
    const editor = (
      await acquireLease({ documentPath, txId: "editor", force: false, leaseMs: 30_000 })
    ).unwrap();
    // Another process is mid-swap: the renewal waits for it.
    await writeFile(lockSwapPathFor(documentPath), "other\n");
    const renewal = editor.renew();
    await sleep(30);
    // The takeover that swap made lands, then the swap ends.
    await writeHolder({ token: "taken-over" });
    await rm(lockSwapPathFor(documentPath));

    const renewed = await renewal;

    expect(renewed.isErr() && renewed.error.code).toBe("locked");
    expect(await tokenOnDisk()).toBe("taken-over");
  });

  test("a forced takeover and a renewal racing always leave the takeover's lock", async () => {
    for (let round = 0; round < 25; round++) {
      const editor = (
        await acquireLease({ documentPath, txId: "editor", force: false, leaseMs: 30_000 })
      ).unwrap();
      if (round % 2 === 0) {
        // Line both up behind one swap so they start together when it ends.
        await writeFile(lockSwapPathFor(documentPath), "other\n");
      }
      const racing = Promise.all([
        editor.renew(),
        acquireLease({ documentPath, txId: "agent", force: true }),
      ]);
      if (round % 2 === 0) {
        await sleep(10);
        await rm(lockSwapPathFor(documentPath));
      }
      const [renewed, forced] = await racing;
      const agent = forced.unwrap();

      expect([round, await tokenOnDisk()]).toEqual([round, agent.holder.token]);
      expect((await editor.verify()).isErr()).toBe(true);
      await editor.release();
      expect([round, await tokenOnDisk()]).toEqual([round, agent.holder.token]);
      expect(renewed.isOk() || renewed.error.code === "locked").toBe(true);
      await agent.release();
      expect(await readdir(dir)).toEqual([]);
    }
  });
});
