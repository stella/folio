/**
 * Render a document's pages to HTML with `folio render`, in a child process.
 * Layout is CPU-bound; running it outside the extension host keeps a large
 * document from stalling every other extension, and lets a stale render be
 * killed when the file changes again.
 *
 * The bytes are copied to a private temporary directory first, so the render
 * works for any file system the editor can read and never touches the folder
 * the document lives in.
 */

import { mkdtemp, readFile, realpath, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import { runCli, type ProcessResult } from "./process";
import type { CliRuntime } from "./runtime";

export type RenderOutcome =
  | { readonly type: "document"; readonly html: string; readonly pageCount: number }
  | { readonly type: "error"; readonly message: string; readonly hint?: string }
  | { readonly type: "cancelled" };

/** What `folio render --output json` printed, read as an outcome. */
export type RenderEnvelope =
  | { readonly type: "ok"; readonly pageCount: number }
  | { readonly type: "error"; readonly message: string; readonly hint?: string };

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

/** The envelope on the last non-empty line of stdout, or `null` when there is none. */
export const parseRenderEnvelope = (stdout: string): RenderEnvelope | null => {
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
    const pageCount = parsed["data"]["pageCount"];
    return typeof pageCount === "number" && Number.isInteger(pageCount) && pageCount >= 0
      ? { type: "ok", pageCount }
      : null;
  }
  if (parsed["ok"] === false && isRecord(parsed["error"])) {
    const { message, hint } = parsed["error"];
    if (typeof message !== "string") return null;
    return typeof hint === "string" ? { type: "error", message, hint } : { type: "error", message };
  }
  return null;
};

/** The last lines of a crashed process's stderr, for the error state. */
export const stderrTail = (stderr: string, lines = 5): string =>
  stderr.trim().split("\n").slice(-lines).join("\n");

export type RenderOptions = {
  readonly runtime: CliRuntime;
  readonly bytes: Uint8Array;
  /** The document's name, for messages. */
  readonly fileName: string;
  readonly signal: AbortSignal;
  readonly timeoutMs?: number;
};

const DEFAULT_TIMEOUT_MS = 120_000;

const describe = (error: unknown): string =>
  error instanceof Error ? error.message : String(error);

/** Render every page of the document; `cancelled` once `signal` aborts. */
export const renderDocument = async ({
  runtime,
  bytes,
  fileName,
  signal,
  timeoutMs = DEFAULT_TIMEOUT_MS,
}: RenderOptions): Promise<RenderOutcome> => {
  // The real path: the CLI reports paths resolved through symlinks.
  const directory = await realpath(await mkdtemp(path.join(tmpdir(), "folio-preview-")));
  try {
    const input = path.join(directory, "document.docx");
    const output = path.join(directory, "document.html");
    await writeFile(input, bytes);
    let result: ProcessResult;
    try {
      result = await runCli({
        runtime,
        args: ["render", input, "-o", output, "--output", "json"],
        signal,
        timeoutMs,
      });
    } catch (error) {
      if (signal.aborted) return { type: "cancelled" };
      return { type: "error", message: `The renderer could not start: ${describe(error)}` };
    }
    if (signal.aborted) return { type: "cancelled" };
    if (result.timedOut) {
      return {
        type: "error",
        message: `${fileName} took longer than ${String(Math.ceil(timeoutMs / 1000))} seconds to render.`,
      };
    }
    const envelope = parseRenderEnvelope(result.stdout);
    if (envelope === null) {
      const tail = stderrTail(result.stderr);
      return {
        type: "error",
        message: `The renderer exited with code ${String(result.code)}.${tail === "" ? "" : `\n${tail}`}`,
      };
    }
    if (envelope.type === "error") {
      // The CLI names the temporary copy; the reader knows the file by its name.
      const message = envelope.message.replaceAll(input, fileName);
      return envelope.hint === undefined
        ? { type: "error", message }
        : { type: "error", message, hint: envelope.hint };
    }
    const html = await readFile(output, "utf8");
    return { type: "document", html, pageCount: envelope.pageCount };
  } finally {
    await rm(directory, { recursive: true, force: true });
  }
};
