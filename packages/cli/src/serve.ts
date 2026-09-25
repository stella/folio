/**
 * `folio serve`: a read-only live preview of one document on 127.0.0.1.
 *
 * The page is the DOM backend's rendering of the document's display list,
 * the same pages `folio render` writes. The server watches the file (and the
 * `.folio/journal.jsonl` a folio write appends to) and tells the page over
 * server-sent events when the file's version changes, so an edit made by the
 * CLI, the MCP server or any other program shows up without a reload.
 *
 * It never writes: not the document, not a lock, not the journal. It binds
 * to 127.0.0.1 only, every URL carries a random token, requests whose `Host`
 * is not the loopback address it bound are refused (so a web page cannot reach
 * it through a rebound name), and only GET and HEAD are answered.
 */

import { Result } from "better-result";
import { randomBytes } from "node:crypto";
import { watch, type FSWatcher } from "node:fs";
import { stat } from "node:fs/promises";
import { createServer, type IncomingMessage, type ServerResponse } from "node:http";
import path from "node:path";

import { escapeHtmlText } from "@stll/folio-core/display-list/html/renderDisplayListToHtml";

import { readDocumentFile } from "./document";
import { journalPathFor } from "./journal";
import { buildDisplayList, displayListHtml } from "./render";
import { inspectPath, readSidecarFile, SIDECAR_DIRECTORY } from "./sidecar";

export const PREVIEW_HOST = "127.0.0.1";

/** How often the file is re-checked when a watch event may have been missed. */
const POLL_INTERVAL_MS = 1000;

export type PreviewServerOptions = {
  documentPath: string;
  /** 0 picks a free port. */
  port: number;
};

export type PreviewServer = {
  /** The preview URL, token included. */
  url: string;
  close: () => Promise<void>;
};

type FileStamp = { size: number; modifiedMs: number; ino: number };

type Rendered =
  | { type: "page"; fileVersion: string; html: string; pageCount: number }
  | { type: "error"; fileVersion: string | null; message: string };

/** The latest `toVersion` the journal committed for this document, if any. */
const journalToVersion = async (documentPath: string): Promise<string | null> => {
  const journalPath = journalPathFor(documentPath, undefined);
  // Read only through a plain `.folio` directory; the preview never creates one.
  const parent = await inspectPath(path.dirname(journalPath));
  if (parent.isErr() || parent.value.type !== "directory") return null;
  const bytes = await readSidecarFile(journalPath);
  if (bytes.isErr()) return null;
  const lines = new TextDecoder().decode(bytes.value).split("\n");
  for (let index = lines.length - 1; index >= 0; index--) {
    const parsed = Result.try((): unknown => JSON.parse(lines[index] ?? ""));
    if (!parsed.isOk() || typeof parsed.value !== "object" || parsed.value === null) continue;
    const entry: Record<string, unknown> = { ...parsed.value };
    if (entry["type"] === "commit" && entry["path"] === documentPath) {
      return typeof entry["toVersion"] === "string" ? entry["toVersion"] : null;
    }
  }
  return null;
};

const SHELL_STYLE =
  "html, body { margin: 0; height: 100%; background: #e8e8e8; } iframe { border: 0; width: 100%; height: 100%; display: block; }";

const shellHtml = (title: string, nonce: string): string =>
  [
    "<!doctype html>",
    `<html><head><meta charset="utf-8"><title>${escapeHtmlText(title)}</title>`,
    `<style nonce="${nonce}">${SHELL_STYLE}</style></head><body>`,
    '<iframe id="document" sandbox src="document.html" title="document"></iframe>',
    `<script nonce="${nonce}">`,
    "const frame = document.getElementById('document');",
    "let version = null;",
    "const events = new EventSource('events');",
    "events.onmessage = (event) => {",
    "  const next = JSON.parse(event.data).fileVersion;",
    "  if (version !== null && next !== version) frame.src = 'document.html?v=' + next;",
    "  version = next;",
    "};",
    "</script></body></html>",
  ].join("\n");

const errorHtml = (message: string): string =>
  `<!doctype html><html><head><meta charset="utf-8"><title>folio</title></head><body><pre>${escapeHtmlText(message)}</pre></body></html>`;

const SECURITY_HEADERS = {
  "Cache-Control": "no-store",
  "Referrer-Policy": "no-referrer",
  "X-Content-Type-Options": "nosniff",
  "Cross-Origin-Resource-Policy": "same-origin",
} as const;

/** The rendered document may contain inline styles and data URLs, and nothing that runs. */
const DOCUMENT_CSP = "default-src 'none'; img-src data:; font-src data:; style-src 'unsafe-inline'";

/** Serve the preview until {@link PreviewServer.close}. */
export const startPreviewServer = async ({
  documentPath,
  port,
}: PreviewServerOptions): Promise<Result<PreviewServer, Error>> => {
  const initial = await readDocumentFile(documentPath);
  if (initial.isErr()) return Result.err(new Error(initial.error.message));
  const realPath = initial.value.path;
  const title = path.basename(realPath);
  const token = randomBytes(24).toString("base64url");
  const prefix = `/${token}/`;

  let stamp: FileStamp | null = null;
  let rendered: Rendered | null = null;
  let rendering: Promise<Rendered> | null = null;
  const listeners = new Set<ServerResponse>();

  /** Render when the file changed since the last render; one render at a time. */
  const current = async (): Promise<Rendered> => {
    const info = await Result.tryPromise(() => stat(realPath));
    const next = info.isOk()
      ? { size: info.value.size, modifiedMs: info.value.mtimeMs, ino: info.value.ino }
      : null;
    const unchanged =
      rendered !== null &&
      next !== null &&
      stamp !== null &&
      next.size === stamp.size &&
      next.modifiedMs === stamp.modifiedMs &&
      next.ino === stamp.ino;
    if (unchanged && rendered !== null) return rendered;
    rendering ??= (async (): Promise<Rendered> => {
      const file = await readDocumentFile(realPath);
      if (file.isErr()) return { type: "error", fileVersion: null, message: file.error.message };
      if (rendered !== null && rendered.fileVersion === file.value.fileVersion) return rendered;
      const list = await buildDisplayList(file.value);
      if (list.isErr()) {
        return { type: "error", fileVersion: file.value.fileVersion, message: list.error.message };
      }
      return {
        type: "page",
        fileVersion: file.value.fileVersion,
        html: displayListHtml(list.value, title),
        pageCount: list.value.pages.length,
      };
    })();
    const result = await rendering;
    rendering = null;
    const changed = rendered === null || rendered.fileVersion !== result.fileVersion;
    rendered = result;
    stamp = next;
    if (changed) {
      const message = `data: ${JSON.stringify({ fileVersion: result.fileVersion })}\n\n`;
      for (const listener of listeners) listener.write(message);
    }
    return result;
  };

  const check = () => {
    void current();
  };

  let boundPort = port;
  const allowedHosts = () =>
    new Set([`${PREVIEW_HOST}:${String(boundPort)}`, `localhost:${String(boundPort)}`]);

  const respond = (
    response: ServerResponse,
    status: number,
    body: string,
    headers: Record<string, string> = {},
  ) => {
    response.writeHead(status, { ...SECURITY_HEADERS, ...headers });
    response.end(body);
  };

  const handle = async (request: IncomingMessage, response: ServerResponse): Promise<void> => {
    if (!allowedHosts().has(request.headers.host ?? "")) {
      respond(response, 403, "forbidden", { "Content-Type": "text/plain" });
      return;
    }
    if (request.method !== "GET" && request.method !== "HEAD") {
      respond(response, 405, "method not allowed", {
        "Content-Type": "text/plain",
        Allow: "GET, HEAD",
      });
      return;
    }
    const url = new URL(request.url ?? "/", `http://${PREVIEW_HOST}`);
    if (!url.pathname.startsWith(prefix)) {
      respond(response, 404, "not found", { "Content-Type": "text/plain" });
      return;
    }
    const route = url.pathname.slice(prefix.length);
    if (route === "") {
      const nonce = randomBytes(16).toString("base64");
      respond(response, 200, shellHtml(title, nonce), {
        "Content-Type": "text/html; charset=utf-8",
        "Content-Security-Policy": `default-src 'none'; script-src 'nonce-${nonce}'; style-src 'nonce-${nonce}'; frame-src 'self'; connect-src 'self'`,
      });
      return;
    }
    if (route === "document.html") {
      const page = await current();
      respond(response, 200, page.type === "page" ? page.html : errorHtml(page.message), {
        "Content-Type": "text/html; charset=utf-8",
        "Content-Security-Policy": DOCUMENT_CSP,
      });
      return;
    }
    if (route === "version") {
      const page = await current();
      respond(
        response,
        200,
        JSON.stringify({
          path: realPath,
          fileVersion: page.fileVersion,
          journalToVersion: await journalToVersion(realPath),
          ...(page.type === "page" ? { pageCount: page.pageCount } : { error: page.message }),
        }),
        { "Content-Type": "application/json" },
      );
      return;
    }
    if (route === "events") {
      response.writeHead(200, {
        ...SECURITY_HEADERS,
        "Content-Type": "text/event-stream",
        Connection: "keep-alive",
      });
      listeners.add(response);
      request.on("close", () => listeners.delete(response));
      const page = await current();
      response.write(`data: ${JSON.stringify({ fileVersion: page.fileVersion })}\n\n`);
      return;
    }
    respond(response, 404, "not found", { "Content-Type": "text/plain" });
  };

  const server = createServer((request, response) => {
    handle(request, response).catch((error: unknown) => {
      respond(response, 500, error instanceof Error ? error.message : "error", {
        "Content-Type": "text/plain",
      });
    });
  });
  const listening = await Result.tryPromise(
    () =>
      new Promise<void>((resolve, reject) => {
        server.once("error", reject);
        server.listen(port, PREVIEW_HOST, () => resolve());
      }),
  );
  if (listening.isErr()) {
    return Result.err(new Error(`Cannot listen on ${PREVIEW_HOST}:${String(port)}.`));
  }
  const address = server.address();
  boundPort = typeof address === "object" && address !== null ? address.port : port;

  // Watch the directory rather than the file: a folio write replaces the file
  // by rename, which ends a watch on the old inode.
  const watchers: FSWatcher[] = [];
  const directory = path.dirname(realPath);
  const watched = new Set([path.basename(realPath), SIDECAR_DIRECTORY, "journal.jsonl"]);
  for (const target of [directory, path.join(directory, SIDECAR_DIRECTORY)]) {
    const watcher = Result.try(() =>
      watch(target, (_event, name) => {
        if (name === null || watched.has(String(name))) check();
      }),
    );
    if (watcher.isOk()) watchers.push(watcher.value);
  }
  const poll = setInterval(check, POLL_INTERVAL_MS);
  check();

  return Result.ok({
    url: `http://${PREVIEW_HOST}:${String(boundPort)}${prefix}`,
    close: async () => {
      clearInterval(poll);
      for (const watcher of watchers) watcher.close();
      for (const listener of listeners) listener.end();
      await new Promise<void>((resolve) => {
        server.close(() => resolve());
        server.closeAllConnections();
      });
    },
  });
};
