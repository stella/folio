import { afterEach, beforeEach, describe, expect, test } from "bun:test";
import { link, mkdir, readdir, readFile, rm, symlink, writeFile } from "node:fs/promises";
import path from "node:path";

import { FolioDocxReviewer } from "@stll/folio-core/server";

import { CONTRACT_PARAGRAPHS, makeTempDir, writeDocx } from "./__tests__/fixtures";
import { fileVersionOf } from "./document";
import { executeWriteTool, type WriteOptions } from "./execute-write";
import { stagePathFor } from "./journal";
import { acquireLease } from "./lock";
import { findFileTool } from "./registry";
import { pruneBackups } from "./transaction";

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
  sourcePrecondition: "required",
  txId: undefined,
  journalPath: undefined,
  mode: "tracked-changes",
  ...overrides,
});

const versionOf = async (filePath: string): Promise<string> =>
  fileVersionOf(new Uint8Array(await readFile(filePath)));

type WriteCall = {
  overrides?: Partial<WriteOptions>;
  /** `current` (the default) reads the source's version now; `none` names none. */
  fileVersion?: string;
  source?: string;
};

const write = async (
  toolName: string,
  args: Record<string, unknown>,
  { overrides = {}, fileVersion = "current", source = file }: WriteCall = {},
) => {
  const tool = findFileTool(toolName);
  if (!tool) throw new Error(`${toolName} is not registered`);
  let version: string | undefined = fileVersion;
  if (fileVersion === "current") version = await versionOf(source);
  if (fileVersion === "none") version = undefined;
  return await executeWriteTool(
    tool,
    { path: source, fileVersion: version, args },
    options(overrides),
  );
};

const REPLACE_FIFTY = {
  operations: [{ type: "replaceInBlock", blockId: "10000002", find: "$50", replace: "$500" }],
};

const reopen = async (filePath: string): Promise<FolioDocxReviewer> => {
  const bytes = new Uint8Array(await readFile(filePath));
  return await FolioDocxReviewer.fromBuffer(bytes.buffer);
};

describe("suggest_changes in place", () => {
  test("commits a selective save with a backup, a journal line, and stamped revisions", async () => {
    const before = await versionOf(file);

    const receipt = (await write("suggest_changes", REPLACE_FIFTY)).unwrap();

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
    const first = (
      await write("suggest_changes", REPLACE_FIFTY, { overrides: { txId: "tx-1" } })
    ).unwrap();
    const after = await versionOf(file);

    const replayed = (
      await write("suggest_changes", REPLACE_FIFTY, { overrides: { txId: "tx-1" } })
    ).unwrap();
    const conflict = await write(
      "suggest_changes",
      { operations: [{ type: "deleteBlock", blockId: "10000004" }] },
      { overrides: { txId: "tx-1" } },
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
      () => write("suggest_changes", REPLACE_FIFTY, { fileVersion: "0".repeat(64) }),
      "stale_version",
    ],
    [
      "no fileVersion",
      () => write("suggest_changes", REPLACE_FIFTY, { fileVersion: "none" }),
      "invalid_input",
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
      () => write("suggest_changes", REPLACE_FIFTY, { overrides: { txId: "../escape" } }),
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
    const forced = await write("suggest_changes", REPLACE_FIFTY, { overrides: { force: true } });

    expect(refused.isErr() && refused.error.code).toBe("locked");
    expect(forced.isOk()).toBe(true);
    await other.release();
  });
});

const toFile = (target: string, extra: { overwrite?: boolean; expectedVersion?: string } = {}) => ({
  overrides: {
    destination: {
      type: "file" as const,
      path: target,
      overwrite: extra.overwrite ?? false,
      expectedVersion: extra.expectedVersion,
    },
  },
});

describe("destinations", () => {
  test("-o never touches the source", async () => {
    const out = path.join(dir, "out.docx");
    const source = await versionOf(file);

    const receipt = (await write("suggest_changes", REPLACE_FIFTY, toFile(out))).unwrap();

    expect(receipt["source"]).toEqual({ path: file, fileVersion: source });
    expect(await versionOf(file)).toBe(source);
    expect((await reopen(out)).getChanges().length).toBe(2);
  });

  test("replacing an existing file needs --overwrite and its version, and backs it up", async () => {
    const out = await writeDocx(dir, "out.docx", [{ text: "Occupied.", paraId: "20000001" }]);
    const occupied = await versionOf(out);

    const refused = await write("suggest_changes", REPLACE_FIFTY, toFile(out));
    const unversioned = await write(
      "suggest_changes",
      REPLACE_FIFTY,
      toFile(out, { overwrite: true }),
    );
    const stale = await write(
      "suggest_changes",
      REPLACE_FIFTY,
      toFile(out, { overwrite: true, expectedVersion: "0".repeat(64) }),
    );
    expect(refused.isErr() && refused.error.code).toBe("destination_exists");
    expect(unversioned.isErr() && unversioned.error.code).toBe("invalid_input");
    expect(stale.isErr() && stale.error.code).toBe("stale_version");
    expect(await versionOf(out)).toBe(occupied);

    const replaced = (
      await write(
        "suggest_changes",
        REPLACE_FIFTY,
        toFile(out, { overwrite: true, expectedVersion: occupied }),
      )
    ).unwrap();

    expect(replaced["backup"]).toBe(
      path.join(dir, ".folio", "backups", "out.docx", `${occupied}.docx`),
    );
    expect(await versionOf(String(replaced["backup"]))).toBe(occupied);
  });

  test("refuses destinations that are not a plain .docx or that alias an input", async () => {
    await writeFile(path.join(dir, "notes.txt"), "keep");
    const revised = await writeDocx(dir, "revised.docx", CONTRACT_PARAGRAPHS);
    await symlink(file, path.join(dir, "alias.docx"));
    await symlink(path.join(dir, "notes.txt"), path.join(dir, "disguised.docx"));
    await symlink(revised, path.join(dir, "revised-alias.docx"));
    await link(file, path.join(dir, "hardlink.docx"));
    await mkdir(path.join(dir, ".folio"));

    const cases: [string, string][] = [
      [path.join(dir, "notes.txt"), "invalid_destination"],
      [path.join(dir, ".hidden.docx"), "invalid_destination"],
      [path.join(dir, ".folio", "journal.docx"), "invalid_destination"],
      [path.join(dir, "alias.docx"), "invalid_destination"],
      [path.join(dir, "hardlink.docx"), "invalid_destination"],
      [path.join(dir, "disguised.docx"), "invalid_destination"],
    ];
    for (const [target, code] of cases) {
      const result = await write(
        "suggest_changes",
        REPLACE_FIFTY,
        toFile(target, { overwrite: true }),
      );
      expect([target, result.isErr() && result.error.code]).toEqual([target, code]);
    }
    const redlineOntoRevised = await write(
      "compare_documents",
      { revisedPath: revised },
      toFile(path.join(dir, "revised-alias.docx"), { overwrite: true }),
    );
    expect(redlineOntoRevised.isErr() && redlineOntoRevised.error.code).toBe("invalid_destination");
    expect(await readFile(path.join(dir, "notes.txt"), "utf8")).toBe("keep");
  });

  test("refuses an in-place write to a file with other hard links", async () => {
    await link(file, path.join(dir, "second-name.docx"));

    const result = await write("suggest_changes", REPLACE_FIFTY);

    expect(result.isErr() && result.error.code).toBe("unsafe_path");
  });

  test("a full repack is reported when allowed", async () => {
    const receipt = (
      await write(
        "suggest_changes",
        { operations: [{ type: "insertAfterBlock", blockId: "10000004", text: "New clause." }] },
        { overrides: { repack: "allow" } },
      )
    ).unwrap();

    expect(receipt["saveStrategy"]).toBe("full-repack");
    expect(receipt["repackReason"]).toBe("structuralChange");
  });
});

describe("sidecar boundary", () => {
  test("refuses a .folio symlink that points outside, writing nothing", async () => {
    const { dir: elsewhere, cleanup: cleanupElsewhere } = await makeTempDir();
    try {
      await symlink(elsewhere, path.join(dir, ".folio"));
      const before = await versionOf(file);

      const result = await write("suggest_changes", REPLACE_FIFTY);

      expect(result.isErr() && result.error.code).toBe("unsafe_path");
      expect(await versionOf(file)).toBe(before);
      expect(await readdir(elsewhere)).toEqual([]);
    } finally {
      await cleanupElsewhere();
    }
  });

  test("keeps backups per document and caps each document separately", async () => {
    const other = await writeDocx(dir, "other.docx", CONTRACT_PARAGRAPHS);
    (await write("suggest_changes", REPLACE_FIFTY)).unwrap();
    (await write("suggest_changes", REPLACE_FIFTY, { source: other })).unwrap();
    (await write("resolve_changes", { action: "accept", all: true })).unwrap();
    const backups = path.join(dir, ".folio", "backups");

    expect((await readdir(backups)).toSorted()).toEqual(["contract.docx", "other.docx"]);
    expect((await readdir(path.join(backups, "contract.docx"))).length).toBe(2);
    expect((await readdir(path.join(backups, "other.docx"))).length).toBe(1);

    (await pruneBackups(path.join(backups, "contract.docx"), 1)).unwrap();
    expect((await readdir(path.join(backups, "contract.docx"))).length).toBe(1);
    expect((await readdir(path.join(backups, "other.docx"))).length).toBe(1);
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
  test("a later write rolls a journaled stage forward, then checks its version against the result", async () => {
    const original = await readFile(file);
    (await write("suggest_changes", REPLACE_FIFTY, { overrides: { txId: "tx-1" } })).unwrap();
    const committed = await readFile(file);
    // Simulate a crash between the journal line and the rename.
    await rm(file);
    await writeFile(file, original);
    await writeFile(stagePathFor(file, "tx-1"), committed);
    const comment = { blockId: "10000003", text: "Confirm the rate." };

    const stale = await write("add_comment", comment);
    const next = (await write("add_comment", comment)).unwrap();

    expect(stale.isErr() && stale.error.code).toBe("stale_version");
    expect(stale.isErr() && JSON.stringify(stale.error.details)).toContain('"rolledForward"');
    expect(next["fromVersion"]).toBe(fileVersionOf(new Uint8Array(committed)));
    const reviewer = await reopen(file);
    expect(reviewer.getChanges().length).toBe(2);
    expect(reviewer.getComments().map(({ text }) => text)).toEqual(["Confirm the rate."]);
  });
});
