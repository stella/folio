/**
 * Drives the `folio` executable as a separate process over a synthetic
 * package: read, suggest a tracked change, list it, accept it, and compare
 * the result with the original.
 */

import { afterAll, beforeAll, describe, expect, test } from "bun:test";
import { spawnSync } from "node:child_process";
import { copyFile, writeFile } from "node:fs/promises";
import path from "node:path";

import { CONTRACT_PARAGRAPHS, makeTempDir, writeDocx } from "./__tests__/fixtures";
import { dataOf, envelopeOf, ISOLATED_GIT_ENV } from "./__tests__/io";

const BIN = path.join(import.meta.dir, "bin.ts");

let dir = "";
let cleanup: () => Promise<void> = () => Promise.resolve();

beforeAll(async () => {
  ({ dir, cleanup } = await makeTempDir());
});

afterAll(async () => {
  await cleanup();
});

const folio = (args: readonly string[], env: Record<string, string | undefined> = {}) => {
  const run = spawnSync(process.execPath, [BIN, ...args], {
    cwd: dir,
    env: { ...ISOLATED_GIT_ENV, FOLIO_AUTHOR: "E2E Reviewer", ...env },
    encoding: "utf8",
  });
  return { exitCode: run.status, stdout: run.stdout, stderr: run.stderr };
};

const resultOf = (data: Record<string, unknown>): unknown => data["result"];

/** Each step starts a process; a loaded machine needs more than the default 5 s. */
const PROCESS_TEST_TIMEOUT_MS = 120_000;

describe("folio executable", () => {
  test(
    "read -> suggest -> changes -> accept --all -> compare",
    async () => {
      const file = await writeDocx(dir, "contract.docx", CONTRACT_PARAGRAPHS);
      const original = path.join(dir, "original.docx");
      await copyFile(file, original);

      const read = folio(["read", file]);
      expect(read.exitCode).toBe(0);
      const readData = dataOf(read.stdout);
      const blocks = resultOf(readData);
      expect(JSON.stringify(blocks)).toContain('"blockId":"10000002"');

      const ops = path.join(dir, "ops.json");
      await writeFile(
        ops,
        JSON.stringify([
          { type: "replaceInBlock", blockId: "10000002", find: "$50", replace: "$500" },
        ]),
      );
      const suggested = folio([
        "suggest",
        file,
        "--input",
        `@${ops}`,
        "--in-place",
        "--expect-version",
        String(readData["fileVersion"]),
        "--date",
        "2026-01-02T03:04:05Z",
      ]);
      expect(suggested.exitCode).toBe(0);
      const receipt = dataOf(suggested.stdout);
      expect(receipt["saveStrategy"]).toBe("selective");

      const stale = folio([
        "suggest",
        file,
        "--input",
        `@${ops}`,
        "--in-place",
        "--expect-version",
        String(readData["fileVersion"]),
      ]);
      expect(stale.exitCode).toBe(10);
      expect(JSON.stringify(envelopeOf(stale.stdout)["error"])).toContain("stale_version");

      const changes = folio(["changes", file]);
      const pending = resultOf(dataOf(changes.stdout));
      expect(Array.isArray(pending) && pending.length).toBe(2);
      expect(JSON.stringify(pending)).toContain('"author":"E2E Reviewer"');

      const unversioned = folio(["accept", file, "--all", "--in-place"]);
      expect(unversioned.exitCode).toBe(2);
      const accepted = folio([
        "accept",
        file,
        "--all",
        "--in-place",
        "--expect-version",
        String(dataOf(changes.stdout)["fileVersion"]),
      ]);
      expect(accepted.exitCode).toBe(0);
      expect(resultOf(dataOf(accepted.stdout))).toEqual({
        action: "accept",
        resolved: 2,
        remaining: 0,
      });
      expect(resultOf(dataOf(folio(["changes", file]).stdout))).toEqual([]);

      const diff = folio(["compare", original, file, "--output", "text"]);
      expect(diff.exitCode).toBe(0);
      expect(diff.stdout).toContain("[- $50-]{+ $500+}");

      const redline = path.join(dir, "redline.docx");
      const written = folio(["compare", original, file, "-o", redline, "--no-expect-version"]);
      expect(written.exitCode).toBe(0);
      expect(dataOf(written.stdout)["saveStrategy"]).toBe("redline");
      const redlineChanges = resultOf(dataOf(folio(["changes", redline]).stdout));
      expect(JSON.stringify(redlineChanges)).toContain("$500");
    },
    PROCESS_TEST_TIMEOUT_MS,
  );

  test(
    "refuses a write with no author configured",
    async () => {
      const file = await writeDocx(dir, "unowned.docx", CONTRACT_PARAGRAPHS);

      const run = folio(
        [
          "comment",
          file,
          "--block-id",
          "10000002",
          "--text",
          "?",
          "--in-place",
          "--no-expect-version",
        ],
        {
          FOLIO_AUTHOR: undefined,
        },
      );

      expect(run.exitCode).toBe(2);
      expect(JSON.stringify(envelopeOf(run.stdout)["error"])).toContain("author_required");
    },
    PROCESS_TEST_TIMEOUT_MS,
  );
});
