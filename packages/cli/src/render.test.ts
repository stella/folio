import { afterEach, beforeEach, describe, expect, test } from "bun:test";
import { Result } from "better-result";
import { existsSync } from "node:fs";
import { readdir, readFile, writeFile } from "node:fs/promises";
import { createRequire } from "node:module";
import path from "node:path";

import { CONTRACT_PARAGRAPHS, makeTempDir, writeDocx } from "./__tests__/fixtures";
import { captureIo, dataOf, envelopeOf } from "./__tests__/io";
import { runFolioCli } from "./cli";

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

const render = async (...args: string[]) => {
  const captured = captureIo();
  const exit = await runFolioCli(["render", file, ...args], captured.io);
  return { exit, stdout: captured.stdout() };
};

/** Whether this machine can run the PNG path: playwright-core and its Chromium. */
const chromiumAvailable = (): boolean => {
  const loaded = Result.try((): unknown => createRequire(import.meta.url)("playwright-core"));
  if (loaded.isErr()) return false;
  const module = loaded.value;
  if (typeof module !== "object" || module === null || !("chromium" in module)) return false;
  const { chromium } = module;
  if (typeof chromium !== "object" || chromium === null || !("executablePath" in chromium)) {
    return false;
  }
  const { executablePath } = chromium;
  if (typeof executablePath !== "function") return false;
  const executable: unknown = Result.try((): unknown => executablePath.call(chromium)).unwrapOr(
    null,
  );
  return typeof executable === "string" && existsSync(executable);
};

const RENDER_TIMEOUT_MS = 120_000;

describe("folio render", () => {
  test(
    "writes a PDF of every page, or one page, and never touches the document",
    async () => {
      const before = await readFile(file);
      const all = await render("-o", path.join(dir, "all.pdf"), "--date", "2026-01-02T03:04:05Z");
      const one = await render("-o", path.join(dir, "one.pdf"), "--page", "1");

      expect(all.exit).toBe(0);
      expect(dataOf(all.stdout)["pageCount"]).toBe(1);
      expect(
        new TextDecoder().decode((await readFile(path.join(dir, "all.pdf"))).slice(0, 5)),
      ).toBe("%PDF-");
      expect(dataOf(one.stdout)["format"]).toBe("pdf");
      expect(Buffer.compare(await readFile(file), before)).toBe(0);
      expect((await readdir(dir)).toSorted()).toEqual(["all.pdf", "contract.docx", "one.pdf"]);
    },
    RENDER_TIMEOUT_MS,
  );

  test(
    "writes HTML pages from the DOM backend",
    async () => {
      const result = await render("-o", path.join(dir, "pages.html"));

      expect(result.exit).toBe(0);
      const html = await readFile(path.join(dir, "pages.html"), "utf8");
      expect(html).toContain('class="layout-page"');
      expect(html).toContain("The buyer pays $50 on signing.");
      expect(html).not.toContain("blob:");
    },
    RENDER_TIMEOUT_MS,
  );

  test(
    "refuses missing pages, unknown formats, and existing outputs",
    async () => {
      await writeFile(path.join(dir, "taken.pdf"), "keep");
      const cases: [string[], string][] = [
        [["-o", path.join(dir, "x.pdf"), "--page", "9"], "invalid_input"],
        [["-o", path.join(dir, "x.docx")], "usage_error"],
        [["-o", path.join(dir, "x.pdf"), "--page", "0"], "usage_error"],
        [[], "usage_error"],
        [["-o", path.join(dir, "taken.pdf")], "destination_exists"],
      ];
      for (const [args, code] of cases) {
        const result = await render(...args);
        expect([args, JSON.stringify(envelopeOf(result.stdout)["error"])]).toEqual([
          args,
          expect.stringContaining(`"code":"${code}"`),
        ]);
      }
      expect(await readFile(path.join(dir, "taken.pdf"), "utf8")).toBe("keep");
    },
    RENDER_TIMEOUT_MS,
  );

  test.skipIf(!chromiumAvailable())(
    "screenshots one page as a PNG through Chromium",
    async () => {
      const result = await render("-o", path.join(dir, "page.png"), "--scale", "1");

      expect(result.exit).toBe(0);
      const png = await readFile(path.join(dir, "page.png"));
      expect([...png.subarray(1, 4)]).toEqual([0x50, 0x4e, 0x47]);
    },
    RENDER_TIMEOUT_MS,
  );
});
