import { describe, expect, test } from "bun:test";

import { decodeBackup } from "./backup";
import type { FolioEditorSaveStrategy, HostMessage } from "./editor-protocol";
import { fileVersionOf, type CliSaveOutcome, type CliSaveRequest } from "./save";
import {
  DocxSession,
  needsRewriteConfirmation,
  SaveCancelled,
  type ConflictChoice,
  type RewriteChoice,
  type RewriteWarning,
  type SessionLease,
  type SessionOptions,
  type StaleChoice,
} from "./session";

const bytesOf = (text: string): Uint8Array => new TextEncoder().encode(text);

const OPENED = bytesOf("PK opened");
const EDITED = bytesOf("PK edited");
const THEIRS = bytesOf("PK theirs");

const SELECTIVE: FolioEditorSaveStrategy = { type: "selective-first" };
const REPACK: FolioEditorSaveStrategy = { type: "full-repack", reason: "structuralChange" };

const settle = () => new Promise((resolve) => setTimeout(resolve, 0));

class FakeLease implements SessionLease {
  token: string | undefined;
  ensured = 0;
  released = 0;
  forgotten = 0;
  disposed = false;

  ensure(): Promise<void> {
    this.ensured += 1;
    this.token = "lease-token";
    return Promise.resolve();
  }

  release(): Promise<void> {
    if (this.token !== undefined) this.released += 1;
    this.token = undefined;
    return Promise.resolve();
  }

  forget(): void {
    this.forgotten += 1;
    this.token = undefined;
  }

  dispose(): Promise<void> {
    this.disposed = true;
    return this.release();
  }
}

type Harness = {
  readonly session: DocxSession;
  readonly posted: HostMessage[];
  readonly saves: CliSaveRequest[];
  readonly lease: FakeLease;
  readonly rewrite: RewriteWarning;
  readonly calls: string[];
  readonly notices: string[];
  disk: Uint8Array;
  /** What the webview answers a `serialize` with. */
  serializeAs: { bytes: Uint8Array; strategy: FolioEditorSaveStrategy };
  /** What the next saves return, in order; then a plain success. */
  outcomes: CliSaveOutcome[];
  rewriteChoice: RewriteChoice | undefined;
  staleChoice: StaleChoice | undefined;
  conflictChoice: ConflictChoice | undefined;
  /** Whether the workbench's save goes through (`false`: the user cancelled a dialog). */
  workbenchSaves: boolean;
  edits: number;
};

const harness = (overrides: Partial<SessionOptions> = {}): Harness => {
  const posted: HostMessage[] = [];
  const saves: CliSaveRequest[] = [];
  const calls: string[] = [];
  const notices: string[] = [];
  const lease = new FakeLease();
  const rewrite: RewriteWarning = {
    enabled: () => true,
    disable: () => {
      calls.push("disableRewriteWarning");
      return Promise.resolve();
    },
    confirmed: false,
  };
  const state = {
    disk: OPENED,
    serializeAs: { bytes: EDITED, strategy: SELECTIVE },
    outcomes: [] as CliSaveOutcome[],
    rewriteChoice: "save" as RewriteChoice | undefined,
    staleChoice: undefined as StaleChoice | undefined,
    conflictChoice: undefined as ConflictChoice | undefined,
    workbenchSaves: true,
    edits: 0,
  };
  let lastLoad = "";
  const session: DocxSession = new DocxSession({
    fileName: "Report.docx",
    documentPath: "/work/Report.docx",
    initial: { bytes: OPENED, baseline: fileVersionOf(OPENED), restored: false },
    author: "Ada Lovelace",
    locale: "en",
    mode: "editing",
    disk: {
      read: () => Promise.resolve(state.disk),
      versionOf: (filePath) =>
        Promise.resolve(filePath === "/work/Existing.docx" ? fileVersionOf(THEIRS) : null),
      write: (filePath) => {
        calls.push(`write ${filePath}`);
        return Promise.resolve();
      },
    },
    saveWithCli: (request) => {
      saves.push(request);
      const next = state.outcomes.shift();
      if (next !== undefined) return Promise.resolve(next);
      if (request.destination === undefined) state.disk = request.bytes;
      return Promise.resolve({
        type: "saved",
        fileVersion: fileVersionOf(request.bytes),
        status: "committed",
      });
    },
    describeChange: (version) => Promise.resolve(`changed to ${version.slice(0, 8)}`),
    lease,
    ui: {
      confirmRewrite: () => {
        calls.push("confirmRewrite");
        return Promise.resolve(state.rewriteChoice);
      },
      resolveStale: () => {
        calls.push("resolveStale");
        return Promise.resolve(state.staleChoice);
      },
      resolveConflict: () => {
        calls.push("resolveConflict");
        return Promise.resolve(state.conflictChoice);
      },
      saveCopy: () => {
        calls.push("saveCopy");
        return Promise.resolve(true);
      },
      saveThroughWorkbench: async () => {
        calls.push("saveThroughWorkbench");
        if (!state.workbenchSaves) return false;
        await session.save();
        return true;
      },
      revertThroughWorkbench: async () => {
        calls.push("revertThroughWorkbench");
        await session.revert();
      },
      notifyUpdated: (message) => notices.push(message),
      showLoadFailed: (message) => calls.push(`loadFailed ${message}`),
      showError: (message) => calls.push(`error ${message}`),
      showLeaseLost: (message) => calls.push(`leaseLost ${message}`),
    },
    rewrite,
    onEdit: () => {
      state.edits += 1;
    },
    ...overrides,
  });
  // The webview: answers `serialize` with the current load's version.
  session.attach((message) => {
    posted.push(message);
    if (message.type === "load" || message.type === "reload") {
      lastLoad = message.document.fileVersion;
    }
    if (message.type === "serialize") {
      const { bytes, strategy } = state.serializeAs;
      // The document the webview has when asked; a reload may overtake the answer.
      const fileVersion = lastLoad;
      queueMicrotask(() =>
        session.handleMessage({
          type: "serialized",
          requestId: message.requestId,
          bytes,
          fileVersion,
          strategy,
        }),
      );
    }
  });
  session.handleMessage({ type: "ready" });
  return Object.assign(state, { session, posted, saves, lease, rewrite, calls, notices });
};

const edit = (h: Harness) => h.session.handleMessage({ type: "edit" });

describe("loading", () => {
  test("answers ready with the document, author, mode, and locale", () => {
    const h = harness();

    const [load] = h.posted;
    expect(load?.type).toBe("load");
    if (load?.type !== "load") return;
    expect(load.document.bytes).toEqual(OPENED);
    expect(load.document.fileName).toBe("Report.docx");
    expect(load.author).toBe("Ada Lovelace");
    expect(load.mode).toBe("editing");
    expect(load.locale).toBe("en");
  });

  test("ignores messages of the wrong shape", () => {
    const h = harness();

    h.session.handleMessage({ type: "edit", extra: 1 });
    h.session.handleMessage({ type: "serialized", requestId: 1 });
    h.session.handleMessage("edit");

    expect(h.edits).toBe(1);
  });
});

describe("edits and undo", () => {
  test("each edit is one workbench edit and takes the lease", async () => {
    const h = harness();

    edit(h);
    edit(h);
    await settle();

    expect(h.edits).toBe(2);
    expect(h.lease.ensured).toBe(2);
    expect(h.session.dirty).toBe(true);
  });

  test("undo and redo reach the webview and track the saved point", () => {
    const h = harness();
    edit(h);

    h.session.undo();
    expect(h.session.dirty).toBe(false);
    h.session.redo();
    expect(h.session.dirty).toBe(true);

    expect(h.posted.slice(1).map((message) => message.type)).toEqual(["undo", "redo"]);
  });
});

describe("save", () => {
  test("saves against the baseline under the lease, then moves the baseline and releases", async () => {
    const h = harness();
    edit(h);
    await settle();

    await h.session.save();

    expect(h.saves).toHaveLength(1);
    expect(h.saves[0]).toMatchObject({
      documentPath: "/work/Report.docx",
      expectedVersion: fileVersionOf(OPENED),
      author: "Ada Lovelace",
      strategy: SELECTIVE,
      leaseToken: "lease-token",
    });
    expect(h.saves[0]?.bytes).toEqual(EDITED);
    expect(h.session.baseline).toBe(fileVersionOf(EDITED));
    expect(h.session.dirty).toBe(false);
    expect(h.lease.released).toBe(1);
  });

  test("keeps the lease when the user typed on during the save", async () => {
    const h = harness();
    edit(h);
    await settle();

    const saving = h.session.save();
    edit(h);
    await saving;

    expect(h.lease.released).toBe(0);
    expect(h.session.dirty).toBe(true);
  });

  test("saves without a token when it holds no lease", async () => {
    const h = harness();

    await h.session.save();

    expect(h.saves[0]?.leaseToken).toBeUndefined();
  });

  test("a stale version asks; Overwrite saves against the version on disk", async () => {
    const h = harness();
    edit(h);
    h.disk = THEIRS;
    h.outcomes = [{ type: "error", code: "stale_version", message: "changed" }];
    h.staleChoice = "overwrite";

    await h.session.save();

    expect(h.calls).toEqual(["resolveStale"]);
    expect(h.saves.map((save) => save.expectedVersion)).toEqual([
      fileVersionOf(OPENED),
      fileVersionOf(THEIRS),
    ]);
    expect(h.session.baseline).toBe(fileVersionOf(EDITED));
  });

  test("a stale version asks; Reload shows the file on disk", async () => {
    const h = harness();
    edit(h);
    h.disk = THEIRS;
    h.outcomes = [{ type: "error", code: "stale_version", message: "changed" }];
    h.staleChoice = "reload";

    await h.session.save();

    const reload = h.posted.at(-1);
    expect(reload?.type).toBe("reload");
    if (reload?.type === "reload") expect(reload.document.bytes).toEqual(THEIRS);
    expect(h.session.baseline).toBe(fileVersionOf(THEIRS));
    expect(h.saves).toHaveLength(1);
  });

  test("a stale version asks; Cancel keeps the document dirty", async () => {
    const h = harness();
    edit(h);
    h.outcomes = [{ type: "error", code: "stale_version", message: "changed" }];
    h.staleChoice = undefined;

    await expect(h.session.save()).rejects.toBeInstanceOf(SaveCancelled);
    expect(h.session.dirty).toBe(true);
    expect(h.session.baseline).toBe(fileVersionOf(OPENED));
  });

  test("a lost lease saves again the way any writer would", async () => {
    const h = harness();
    edit(h);
    await settle();
    h.outcomes = [{ type: "error", code: "locked", message: "not held under that lease token" }];

    await h.session.save();

    expect(h.lease.forgotten).toBe(1);
    expect(h.saves.map((save) => save.leaseToken)).toEqual(["lease-token", undefined]);
  });

  test("any other refusal fails the save with the CLI's message and hint", async () => {
    const h = harness();
    h.outcomes = [{ type: "error", code: "locked", message: "Held by folio-cli.", hint: "Retry." }];

    await expect(h.session.save()).rejects.toThrow("Held by folio-cli. Retry.");
  });

  test("a document off disk cannot be saved in place", async () => {
    const h = harness({ documentPath: null });

    await expect(h.session.save()).rejects.toThrow("read-only");
  });

  test("a serialization of an earlier load is refused", async () => {
    const h = harness();
    const pending = h.session.serialize();
    // A reload lands before the webview answers.
    await h.session.reloadFromDisk(THEIRS);

    await expect(pending).rejects.toThrow("reloaded while it was being saved");
  });
});

describe("the rewrite warning", () => {
  test("is due for a full repack only, while enabled and unanswered", () => {
    const warning: RewriteWarning = {
      enabled: () => true,
      disable: () => Promise.resolve(),
      confirmed: false,
    };

    expect(needsRewriteConfirmation(REPACK, warning)).toBe(true);
    expect(needsRewriteConfirmation(SELECTIVE, warning)).toBe(false);
    expect(needsRewriteConfirmation(REPACK, { ...warning, confirmed: true })).toBe(false);
    expect(needsRewriteConfirmation(REPACK, { ...warning, enabled: () => false })).toBe(false);
  });

  test("asks once per session", async () => {
    const h = harness();
    h.serializeAs = { bytes: EDITED, strategy: REPACK };

    await h.session.save();
    await h.session.save();

    expect(h.calls).toEqual(["confirmRewrite"]);
    expect(h.saves).toHaveLength(2);
    expect(h.saves[0]?.strategy).toEqual(REPACK);
  });

  test("Always turns the setting off", async () => {
    const h = harness();
    h.serializeAs = { bytes: EDITED, strategy: REPACK };
    h.rewriteChoice = "always";

    await h.session.save();

    expect(h.calls).toEqual(["confirmRewrite", "disableRewriteWarning"]);
    expect(h.saves).toHaveLength(1);
  });

  test("Save a Copy saves a copy and leaves the document unsaved", async () => {
    const h = harness();
    h.serializeAs = { bytes: EDITED, strategy: REPACK };
    h.rewriteChoice = "saveCopy";

    await expect(h.session.save()).rejects.toBeInstanceOf(SaveCancelled);

    expect(h.calls).toEqual(["confirmRewrite", "saveCopy"]);
    expect(h.saves).toHaveLength(0);
    expect(h.rewrite.confirmed).toBe(false);
  });

  test("Cancel saves nothing", async () => {
    const h = harness();
    h.serializeAs = { bytes: EDITED, strategy: REPACK };
    h.rewriteChoice = undefined;

    await expect(h.session.save()).rejects.toBeInstanceOf(SaveCancelled);
    expect(h.saves).toHaveLength(0);
  });
});

describe("save as", () => {
  test("saves a copy with folio save -o, leaving the document's baseline alone", async () => {
    const h = harness();

    await h.session.saveCopyTo("/work/Copy.docx");

    expect(h.saves[0]).toMatchObject({
      documentPath: "/work/Report.docx",
      expectedVersion: fileVersionOf(OPENED),
      destination: { path: "/work/Copy.docx", expectedVersion: null },
    });
    expect(h.session.baseline).toBe(fileVersionOf(OPENED));
  });

  test("replaces an existing file against its current version", async () => {
    const h = harness();

    await h.session.saveCopyTo("/work/Existing.docx");

    expect(h.saves[0]?.destination).toEqual({
      path: "/work/Existing.docx",
      expectedVersion: fileVersionOf(THEIRS),
    });
  });

  test("writes the copy directly when the document is not on disk", async () => {
    const h = harness({ documentPath: null });

    await h.session.saveCopyTo("/work/Copy.docx");

    expect(h.saves).toHaveLength(0);
    expect(h.calls).toEqual(["write /work/Copy.docx"]);
  });
});

describe("changes on disk", () => {
  test("ignores the version it already has", async () => {
    const h = harness();

    await h.session.diskChanged();

    expect(h.posted.map((message) => message.type)).toEqual(["load"]);
  });

  test("a clean document reloads and says who changed it", async () => {
    const h = harness();
    h.disk = THEIRS;

    await h.session.diskChanged();

    expect(h.posted.at(-1)?.type).toBe("reload");
    expect(h.session.baseline).toBe(fileVersionOf(THEIRS));
    expect(h.notices).toEqual([`changed to ${fileVersionOf(THEIRS).slice(0, 8)}`]);
  });

  test("a dirty document asks; Keep Mine makes the next save replace theirs", async () => {
    const h = harness();
    edit(h);
    h.disk = THEIRS;
    h.conflictChoice = "keep";

    await h.session.diskChanged();
    await h.session.save();

    expect(h.calls).toEqual(["resolveConflict"]);
    expect(h.saves[0]?.expectedVersion).toBe(fileVersionOf(THEIRS));
  });

  test("a dirty document asks; Reload Theirs reverts through the workbench", async () => {
    const h = harness();
    edit(h);
    h.disk = THEIRS;
    h.conflictChoice = "reload";

    await h.session.diskChanged();

    expect(h.calls).toEqual(["resolveConflict", "revertThroughWorkbench"]);
    expect(h.session.baseline).toBe(fileVersionOf(THEIRS));
  });

  test("a dirty document asks; Save Mine as Copy saves a copy, then reverts", async () => {
    const h = harness();
    edit(h);
    h.disk = THEIRS;
    h.conflictChoice = "saveCopy";

    await h.session.diskChanged();

    expect(h.calls).toEqual(["resolveConflict", "saveCopy", "revertThroughWorkbench"]);
  });

  test("its own save's rename is not a change", async () => {
    const h = harness();
    edit(h);

    const saving = h.session.save();
    // The watcher fires while the CLI is still running.
    const changed = h.session.diskChanged();
    await Promise.all([saving, changed]);
    await settle();

    expect(h.calls).toEqual([]);
    expect(h.notices).toEqual([]);
    expect(h.posted.map((message) => message.type)).toEqual(["load", "serialize"]);
  });
});

describe("flush requests", () => {
  test("save through the workbench, then let the lease go", async () => {
    const h = harness();
    edit(h);
    await settle();
    // Typing between the save's start and its end would normally keep the lease.
    h.serializeAs = { bytes: EDITED, strategy: SELECTIVE };

    await h.session.flush();

    expect(h.calls).toEqual(["saveThroughWorkbench"]);
    expect(h.saves[0]?.leaseToken).toBe("lease-token");
    expect(h.lease.token).toBeUndefined();
    expect(h.lease.released).toBe(1);
  });

  test("keep the lease when the save was cancelled", async () => {
    const h = harness();
    h.workbenchSaves = false;
    edit(h);
    await settle();

    await h.session.flush();

    expect(h.lease.token).toBe("lease-token");
  });
});

describe("backup and restore", () => {
  test("the backup carries the baseline and the unsaved bytes", async () => {
    const h = harness();
    edit(h);

    const backup = decodeBackup(await h.session.backup());

    expect(backup).toEqual({ baseline: fileVersionOf(OPENED), bytes: EDITED });
  });

  test("a restored document is dirty, holds the lease, and saves against its baseline", async () => {
    const h = harness({
      initial: { bytes: EDITED, baseline: fileVersionOf(OPENED), restored: true },
    });
    await settle();

    expect(h.session.dirty).toBe(true);
    expect(h.lease.ensured).toBe(1);
    await h.session.save();
    expect(h.saves[0]?.expectedVersion).toBe(fileVersionOf(OPENED));
    expect(h.session.dirty).toBe(false);
  });

  test("revert drops the edits and the lease", async () => {
    const h = harness();
    edit(h);
    await settle();

    await h.session.revert();

    expect(h.session.dirty).toBe(false);
    expect(h.lease.released).toBe(1);
    expect(h.posted.at(-1)?.type).toBe("reload");
  });
});
