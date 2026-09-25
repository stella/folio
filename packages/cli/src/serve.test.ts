import { afterEach, beforeEach, describe, expect, test } from "bun:test";
import { readdir, writeFile } from "node:fs/promises";
import { request } from "node:http";

import { buildDocx, CONTRACT_PARAGRAPHS, makeTempDir, writeDocx } from "./__tests__/fixtures";
import { startPreviewServer, type PreviewServer } from "./serve";

let dir = "";
let cleanup: () => Promise<void> = () => Promise.resolve();
let file = "";
let server: PreviewServer | undefined;

beforeEach(async () => {
  ({ dir, cleanup } = await makeTempDir());
  file = await writeDocx(dir, "contract.docx", CONTRACT_PARAGRAPHS);
});

afterEach(async () => {
  await server?.close();
  server = undefined;
  await cleanup();
});

const start = async (): Promise<PreviewServer> => {
  server = (await startPreviewServer({ documentPath: file, port: 0 })).unwrap();
  return server;
};

type Reply = { status: number; body: string; headers: Record<string, unknown> };

/** A raw request, so the test controls the method and the Host header. */
const send = (url: string, options: { method?: string; host?: string } = {}): Promise<Reply> =>
  new Promise((resolve, reject) => {
    const target = new URL(url);
    const outgoing = request(
      {
        hostname: target.hostname,
        port: target.port,
        path: `${target.pathname}${target.search}`,
        method: options.method ?? "GET",
        headers: { host: options.host ?? target.host },
      },
      (response) => {
        let body = "";
        response.setEncoding("utf8");
        response.on("data", (chunk: string) => {
          body += chunk;
        });
        response.on("end", () =>
          resolve({ status: response.statusCode ?? 0, body, headers: response.headers }),
        );
      },
    );
    outgoing.on("error", reject);
    outgoing.end();
  });

const versionOf = async (url: string): Promise<unknown> => {
  const parsed: unknown = JSON.parse((await send(`${url}version`)).body);
  return typeof parsed === "object" && parsed !== null && "fileVersion" in parsed
    ? parsed.fileVersion
    : undefined;
};

const SERVE_TIMEOUT_MS = 120_000;

describe("folio serve", () => {
  test(
    "serves the rendered pages behind a token on 127.0.0.1",
    async () => {
      const { url } = await start();

      const shell = await send(url);
      const page = await send(`${url}document.html`);
      const withoutToken = await send(url.replace(/\/[^/]+\/$/u, "/"));
      const wrongToken = await send(url.replace(/\/[^/]+\/$/u, "/not-the-token/"));

      expect(new URL(url).hostname).toBe("127.0.0.1");
      expect(shell.status).toBe(200);
      expect(shell.body).toContain('<iframe id="document" sandbox');
      expect(String(shell.headers["content-security-policy"])).toContain("script-src 'nonce-");
      expect(page.body).toContain("The buyer pays $50 on signing.");
      expect(String(page.headers["content-security-policy"])).toContain("default-src 'none'");
      expect(page.headers["referrer-policy"]).toBe("no-referrer");
      expect([withoutToken.status, wrongToken.status]).toEqual([404, 404]);
    },
    SERVE_TIMEOUT_MS,
  );

  test(
    "refuses other hosts and methods, and never writes",
    async () => {
      const { url } = await start();

      const rebound = await send(url, { host: `attacker.example:${new URL(url).port}` });
      const posted = await send(url, { method: "POST" });

      expect(rebound.status).toBe(403);
      expect(posted.status).toBe(405);
      expect(await readdir(dir)).toEqual(["contract.docx"]);
    },
    SERVE_TIMEOUT_MS,
  );

  test(
    "re-renders when the file's version changes",
    async () => {
      const { url } = await start();
      const before = await versionOf(url);

      await writeFile(file, await buildDocx([{ text: "Rewritten clause.", paraId: "20000001" }]));
      let after = before;
      for (let attempt = 0; attempt < 50 && after === before; attempt++) {
        // oxlint-disable-next-line no-await-in-loop -- polling for the watcher to see the write
        await new Promise((resolve) => setTimeout(resolve, 100));
        // oxlint-disable-next-line no-await-in-loop -- polling for the watcher to see the write
        after = await versionOf(url);
      }

      expect(after).not.toBe(before);
      expect((await send(`${url}document.html`)).body).toContain("Rewritten clause.");
    },
    SERVE_TIMEOUT_MS,
  );
});
