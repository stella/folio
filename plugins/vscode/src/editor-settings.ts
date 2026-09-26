/**
 * What the editor opens a document with: the author its tracked changes and
 * comments record, and the mode it starts in. Nothing here imports `vscode`.
 */

import { execFile } from "node:child_process";
import { userInfo } from "node:os";

import type { FolioEditingMode } from "./editor-protocol";
import { configuredAuthor } from "./mcp";

export const TRACK_CHANGES_SETTING = "folio.editor.trackChanges";
export const CONFIRM_REWRITE_SETTING = "folio.editor.confirmRewrite";

/** The name when neither the setting, git, nor the OS gives one. */
const FALLBACK_AUTHOR = "Folio user";

/** `folio.author`, else git's `user.name`, else the OS account name. */
export const resolveEditorAuthor = ({
  setting,
  gitUserName,
  osUserName,
}: {
  readonly setting: string | undefined;
  readonly gitUserName: string | undefined;
  readonly osUserName: string | undefined;
}): string =>
  configuredAuthor(setting) ??
  configuredAuthor(gitUserName) ??
  configuredAuthor(osUserName) ??
  FALLBACK_AUTHOR;

/** `git config user.name` as seen from `cwd`, or `undefined`. */
export const readGitUserName = (cwd: string | undefined): Promise<string | undefined> =>
  new Promise((resolve) => {
    execFile(
      "git",
      ["config", "user.name"],
      { ...(cwd !== undefined && { cwd }), timeout: 2000, windowsHide: true },
      (error, stdout) => resolve(error === null ? stdout.trim() : undefined),
    );
  });

export const osUserName = (): string | undefined => {
  try {
    return userInfo().username;
  } catch {
    return undefined;
  }
};

/**
 * `editing` writes directly and `suggesting` records tracked changes, from
 * `folio.editor.trackChanges`. A file the CLI cannot save opens read-only.
 */
export const initialMode = (trackChanges: unknown, saveable: boolean): FolioEditingMode => {
  if (!saveable) return "viewing";
  return trackChanges === "on" ? "suggesting" : "editing";
};
