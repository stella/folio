import { afterAll, describe, expect, test } from "bun:test";
import { existsSync, mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";

import { acquireLeaseForWrite } from "../../../packages/cli/src/editor-lease";
import { lockPathFor, readLease } from "../../../packages/cli/src/lock";
import { EditorLease, type HeldLease, type LeaseBackend } from "./lease";

/** A backend that records what the controller does with it. */
const fakeBackend = () => {
  const log: string[] = [];
  let refuse: string | undefined;
  let lost: ((message: string) => void) | undefined;
  let request: (() => void) | undefined;
  let acquired = 0;
  const backend: LeaseBackend = {
    acquire: async () => {
      await Promise.resolve();
      if (refuse !== undefined) {
        log.push("refused");
        return { ok: false, message: refuse };
      }
      acquired += 1;
      const token = `token-${String(acquired)}`;
      log.push(`acquire ${token}`);
      const lease: HeldLease = {
        token,
        release: () => {
          log.push(`release ${token}`);
          return Promise.resolve();
        },
      };
      return { ok: true, lease };
    },
    keepAlive: (lease, onLost) => {
      log.push(`keepAlive ${lease.token}`);
      lost = onLost;
      return { stop: () => log.push(`stop ${lease.token}`) };
    },
    watch: (_documentPath, token, onRequest) => {
      log.push(`watch ${token}`);
      request = () =>
        onRequest({
          id: "r1",
          leaseToken: token,
          owner: "folio-cli",
          pid: 1,
          host: "h",
          txId: "t",
          requestedAt: "",
          deadline: "",
        });
      return { close: () => log.push(`close ${token}`) };
    },
  };
  return {
    backend,
    log,
    refuseWith: (message: string | undefined) => {
      refuse = message;
    },
    loseLease: (message: string) => lost?.(message),
    requestFlush: () => request?.(),
  };
};

const events = () => {
  const seen: string[] = [];
  return {
    seen,
    handlers: {
      onFlushRequest: () => seen.push("flush"),
      onLost: (message: string) => seen.push(`lost ${message}`),
    },
  };
};

describe("EditorLease", () => {
  test("takes the lease once, renews it, and watches for requests under its token", async () => {
    const fake = fakeBackend();
    const { seen, handlers } = events();
    const lease = new EditorLease("/work/Report.docx", handlers, fake.backend);

    await Promise.all([lease.ensure(), lease.ensure()]);
    await lease.ensure();
    fake.requestFlush();

    expect(lease.token).toBe("token-1");
    expect(fake.log).toEqual(["acquire token-1", "keepAlive token-1", "watch token-1"]);
    expect(seen).toEqual(["flush"]);
  });

  test("a refusal leaves it free, and the next edit tries again", async () => {
    const fake = fakeBackend();
    const lease = new EditorLease("/work/Report.docx", events().handlers, fake.backend);

    fake.refuseWith("A write is waiting to take the lease.");
    await lease.ensure();
    expect(lease.token).toBeUndefined();
    expect(lease.lastRefusal).toBe("A write is waiting to take the lease.");

    fake.refuseWith(undefined);
    await lease.ensure();
    expect(lease.token).toBe("token-1");
    expect(lease.lastRefusal).toBeUndefined();
  });

  test("release stops renewing and watching, then lets the lock go", async () => {
    const fake = fakeBackend();
    const lease = new EditorLease("/work/Report.docx", events().handlers, fake.backend);
    await lease.ensure();

    await lease.release();
    await lease.release();

    expect(lease.token).toBeUndefined();
    expect(fake.log.slice(3)).toEqual(["stop token-1", "close token-1", "release token-1"]);
  });

  test("a release during acquisition lets the new lease go", async () => {
    const fake = fakeBackend();
    const lease = new EditorLease("/work/Report.docx", events().handlers, fake.backend);

    const acquiring = lease.ensure();
    await lease.release();
    await acquiring;

    expect(lease.token).toBeUndefined();
    expect(fake.log.at(-1)).toBe("release token-1");
  });

  test("a lost lease stops, reports, and leaves the lock alone", async () => {
    const fake = fakeBackend();
    const { seen, handlers } = events();
    const lease = new EditorLease("/work/Report.docx", handlers, fake.backend);
    await lease.ensure();

    fake.loseLease("taken over");

    expect(lease.token).toBeUndefined();
    expect(seen).toEqual(["lost taken over"]);
    expect(fake.log).not.toContain("release token-1");
  });

  test("once disposed it takes no lease", async () => {
    const fake = fakeBackend();
    const lease = new EditorLease("/work/Report.docx", events().handlers, fake.backend);

    await lease.dispose();
    await lease.ensure();

    expect(fake.log).toEqual([]);
  });
});

describe("EditorLease with the CLI's lease module", () => {
  const directory = mkdtempSync(path.join(tmpdir(), "folio-lease-test-"));
  afterAll(() => rmSync(directory, { recursive: true, force: true }));

  test("holds an acceptsFlush lease, and a write's flush request reaches it", async () => {
    const documentPath = path.join(directory, "Report.docx");
    writeFileSync(documentPath, "PK");
    const lease: EditorLease = new EditorLease(documentPath, {
      // The editor would save here; then it lets the lease go.
      onFlushRequest: () => void lease.release(),
      onLost: () => undefined,
    });

    await lease.ensure();
    const held = await readLease(documentPath);
    expect(held.type).toBe("held");
    if (held.type === "held") {
      expect(held.holder.owner).toBe("folio-vscode");
      expect(held.holder.acceptsFlush).toBe(true);
      expect(held.holder.pid).toBe(process.pid);
      expect(held.holder.token).toBe(lease.token ?? "");
    }

    const write = await acquireLeaseForWrite({
      documentPath,
      txId: "agent-write",
      force: false,
      flushWaitMs: 5000,
    });

    expect(write.isOk()).toBe(true);
    if (write.isOk()) {
      expect(write.value.flush.type).toBe("flushed");
      await write.value.lease.release();
    }
    expect(lease.token).toBeUndefined();
    expect(existsSync(lockPathFor(documentPath))).toBe(false);
    await lease.dispose();
  });
});
