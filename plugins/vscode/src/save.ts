/**
 * Save an editor's serialized package with `folio save`, in a child process:
 * the CLI commits it as one transaction (lease, version check, backup in
 * `.folio/backups`, journal line) and renames it into place. The bytes go
 * through a private temporary file.
 */

import { createHash } from "node:crypto";
import { mkdtemp, realpath, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import { saveStrategyFlag, type FolioEditorSaveStrategy } from "./editor-protocol";
import { runCli } from "./process";
import { stderrTail } from "./render";
import type { CliRuntime } from "./runtime";

/** The lease owner the editor records, and the surface its saves journal. */
export const EDITOR_LEASE_OWNER = "folio-vscode";
export const EDITOR_SURFACE = "vscode";

/** The folio fileVersion of some bytes: their SHA-256, in hex. */
export const fileVersionOf = (bytes: Uint8Array): string =>
  createHash("sha256").update(bytes).digest("hex");

export type CliSaveRequest = {
  /** The document the editor opened, on disk. */
  readonly documentPath: string;
  readonly bytes: Uint8Array;
  /** The fileVersion the edits were made from. */
  readonly expectedVersion: string;
  readonly author: string;
  readonly strategy: FolioEditorSaveStrategy;
  /** Save under the editor's own lease. */
  readonly leaseToken?: string;
  /** Save to another file (Save As); `expectedVersion` is that file's version when it exists. */
  readonly destination?: { readonly path: string; readonly expectedVersion: string | null };
};

export type CliSaveOutcome =
  | {
      readonly type: "saved";
      /** The file's version now. */
      readonly fileVersion: string;
      /** `committed`, or `unchanged` when the bytes were already on disk. */
      readonly status: string;
      /** Where the previous version was backed up, when it was. */
      readonly backup?: string;
    }
  | {
      readonly type: "error";
      /** The CLI's error code (`stale_version`, `locked`, ...), or `crashed`. */
      readonly code: string;
      readonly message: string;
      readonly hint?: string;
    };

/** `folio save`'s arguments for a request whose bytes are at `fromPath`. */
export const saveArgs = (request: CliSaveRequest, fromPath: string): string[] => [
  "save",
  request.documentPath,
  "--from",
  fromPath,
  "--expect-version",
  request.expectedVersion,
  "--author",
  request.author,
  "--owner",
  EDITOR_LEASE_OWNER,
  "--surface",
  EDITOR_SURFACE,
  "--save-strategy",
  saveStrategyFlag(request.strategy),
  ...(request.leaseToken === undefined ? [] : ["--lease-token", request.leaseToken]),
  ...(request.destination === undefined
    ? []
    : [
        "-o",
        request.destination.path,
        ...(request.destination.expectedVersion === null
          ? []
          : ["--overwrite", "--expect-destination-version", request.destination.expectedVersion]),
      ]),
  "--output",
  "json",
];

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

/** The outcome on the last non-empty line of stdout, or `null` when there is none. */
export const parseSaveEnvelope = (stdout: string): CliSaveOutcome | null => {
  const line = stdout
    .split("\n")
    .map((entry) => entry.trim())
    .findLast((entry) => entry !== "");
  if (line === undefined) return null;
  let parsed: unknown;
  try {
    parsed = JSON.parse(line);
  } catch {
    return null;
  }
  if (!isRecord(parsed)) return null;
  if (parsed["ok"] === true && isRecord(parsed["data"])) {
    const { fileVersion, status, backup } = parsed["data"];
    if (typeof fileVersion !== "string" || typeof status !== "string") return null;
    return typeof backup === "string"
      ? { type: "saved", fileVersion, status, backup }
      : { type: "saved", fileVersion, status };
  }
  if (parsed["ok"] === false && isRecord(parsed["error"])) {
    const { code, message, hint } = parsed["error"];
    if (typeof code !== "string" || typeof message !== "string") return null;
    return typeof hint === "string"
      ? { type: "error", code, message, hint }
      : { type: "error", code, message };
  }
  return null;
};

const DEFAULT_TIMEOUT_MS = 120_000;

const describe = (error: unknown): string =>
  error instanceof Error ? error.message : String(error);

/** Run `folio save` for `request`. Never throws; a failure is an `error` outcome. */
export const saveWithCli = async (
  runtime: CliRuntime,
  request: CliSaveRequest,
  timeoutMs = DEFAULT_TIMEOUT_MS,
): Promise<CliSaveOutcome> => {
  const directory = await realpath(await mkdtemp(path.join(tmpdir(), "folio-save-")));
  try {
    const from = path.join(directory, "saved.docx");
    await writeFile(from, request.bytes);
    const result = await runCli({ runtime, args: saveArgs(request, from), timeoutMs });
    if (result.timedOut) {
      return {
        type: "error",
        code: "crashed",
        message: `folio save took longer than ${String(Math.ceil(timeoutMs / 1000))} seconds.`,
      };
    }
    const outcome = parseSaveEnvelope(result.stdout);
    if (outcome !== null) return outcome;
    const tail = stderrTail(result.stderr);
    return {
      type: "error",
      code: "crashed",
      message: `folio save exited with code ${String(result.code)}.${tail === "" ? "" : `\n${tail}`}`,
    };
  } catch (error) {
    return {
      type: "error",
      code: "crashed",
      message: `folio save could not run: ${describe(error)}`,
    };
  } finally {
    await rm(directory, { recursive: true, force: true });
  }
};
