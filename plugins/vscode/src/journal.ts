/**
 * Who changed a document on disk, from folio's journal
 * (`.folio/journal.jsonl` beside it): every folio write appends a `commit`
 * line naming its tool, author and resulting version. The editor reads it to
 * say "Updated by <author> (<tool>)" when it reloads a file an agent changed.
 */

import { readFile } from "node:fs/promises";
import path from "node:path";

export type JournalCommit = {
  readonly tool: string;
  readonly author: string;
  readonly time: string;
  readonly toVersion: string;
};

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

/**
 * The newest commit in `journal` (the file's text) that produced `version`,
 * and that recovery did not discard. Matching the version rather than the
 * path keeps a symlinked folder from hiding the line.
 */
export const commitForVersion = (journal: string, version: string): JournalCommit | null => {
  const lines = journal.split("\n");
  const discarded = new Set<unknown>();
  for (let index = lines.length - 1; index >= 0; index--) {
    const line = lines[index]?.trim() ?? "";
    if (line === "") continue;
    let entry: unknown;
    try {
      entry = JSON.parse(line);
    } catch {
      continue;
    }
    if (!isRecord(entry)) continue;
    if (entry["type"] === "recovery" && entry["action"] === "discarded") {
      discarded.add(entry["txId"]);
      continue;
    }
    if (entry["type"] !== "commit" || entry["toVersion"] !== version) continue;
    if (discarded.has(entry["txId"])) continue;
    const { tool, author, time } = entry;
    if (typeof tool !== "string" || typeof author !== "string" || typeof time !== "string") {
      return null;
    }
    return { tool, author, time, toVersion: version };
  }
  return null;
};

/** The commit behind `version` of the document at `documentPath`, if the journal has it. */
export const readCommitForVersion = async (
  documentPath: string,
  version: string,
): Promise<JournalCommit | null> => {
  try {
    const journal = await readFile(
      path.join(path.dirname(documentPath), ".folio", "journal.jsonl"),
      "utf8",
    );
    return commitForVersion(journal, version);
  } catch {
    return null;
  }
};

/** A tool name as a reader says it: `add_comment` is "add comment". */
const toolLabel = (tool: string): string =>
  tool === "editor_save" ? "saved in an editor" : tool.replaceAll("_", " ");

/** The notice for a reload: "Updated by Ada (add comment)", or a plain one without a commit. */
export const updatedNotice = (fileName: string, commit: JournalCommit | null): string =>
  commit === null
    ? `${fileName} changed on disk; reloaded.`
    : `${fileName} updated by ${commit.author} (${toolLabel(commit.tool)}).`;
