import { afterAll, describe, expect, test } from "bun:test";
import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";

import { commitForVersion, readCommitForVersion, updatedNotice } from "./journal";

const V1 = "1".repeat(64);
const V2 = "2".repeat(64);

const commit = (txId: string, toVersion: string, tool: string, author: string) =>
  JSON.stringify({ type: "commit", txId, tool, author, time: "2026-09-26T10:00:00Z", toVersion });

describe("commitForVersion", () => {
  test("finds the newest commit that produced the version", () => {
    const journal = [
      commit("a", V1, "add_comment", "Grace"),
      "not json",
      commit("b", V2, "replace_text", "Agent"),
      commit("c", V2, "add_comment", "Agent Two"),
      "",
    ].join("\n");

    expect(commitForVersion(journal, V2)).toEqual({
      tool: "add_comment",
      author: "Agent Two",
      time: "2026-09-26T10:00:00Z",
      toVersion: V2,
    });
    expect(commitForVersion(journal, V1)?.author).toBe("Grace");
    expect(commitForVersion(journal, "9".repeat(64))).toBeNull();
  });

  test("skips a commit recovery discarded", () => {
    const journal = [
      commit("a", V1, "add_comment", "Grace"),
      commit("b", V1, "replace_text", "Agent"),
      JSON.stringify({ type: "recovery", txId: "b", action: "discarded", reason: "stageChanged" }),
    ].join("\n");

    expect(commitForVersion(journal, V1)?.author).toBe("Grace");
  });
});

describe("readCommitForVersion", () => {
  const directory = mkdtempSync(path.join(tmpdir(), "folio-journal-test-"));
  afterAll(() => rmSync(directory, { recursive: true, force: true }));

  test("reads .folio/journal.jsonl beside the document", async () => {
    mkdirSync(path.join(directory, ".folio"));
    writeFileSync(
      path.join(directory, ".folio", "journal.jsonl"),
      `${commit("a", V1, "add_comment", "Grace")}\n`,
    );

    expect((await readCommitForVersion(path.join(directory, "R.docx"), V1))?.author).toBe("Grace");
    expect(await readCommitForVersion(path.join(directory, "missing", "R.docx"), V1)).toBeNull();
  });
});

describe("updatedNotice", () => {
  test("names the author and the tool", () => {
    expect(
      updatedNotice("R.docx", { tool: "add_comment", author: "Agent", time: "", toVersion: V1 }),
    ).toBe("R.docx updated by Agent (add comment).");
    expect(
      updatedNotice("R.docx", { tool: "editor_save", author: "Grace", time: "", toVersion: V1 }),
    ).toBe("R.docx updated by Grace (saved in an editor).");
    expect(updatedNotice("R.docx", null)).toBe("R.docx changed on disk; reloaded.");
  });
});
