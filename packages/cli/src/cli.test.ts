import { afterEach, beforeEach, describe, expect, test } from "bun:test";

import { CONTRACT_PARAGRAPHS, makeTempDir, writeDocx } from "./__tests__/fixtures";
import { captureIo, envelopeOf } from "./__tests__/io";
import { runFolioCli } from "./cli";

let dir = "";
let cleanup: () => Promise<void> = () => Promise.resolve();

beforeEach(async () => {
  ({ dir, cleanup } = await makeTempDir());
});

afterEach(async () => {
  await cleanup();
});

describe("runFolioCli", () => {
  test("prints a JSON envelope off a terminal and text on one", async () => {
    const file = await writeDocx(dir, "contract.docx", CONTRACT_PARAGRAPHS);
    const piped = captureIo();
    const terminal = captureIo({ isTTY: true });

    const pipedExit = await runFolioCli(["read", file], piped.io);
    const terminalExit = await runFolioCli(["read", file], terminal.io);

    expect([pipedExit, terminalExit]).toEqual([0, 0]);
    expect(envelopeOf(piped.stdout())["ok"]).toBe(true);
    expect(terminal.stdout()).toContain("[10000002] The buyer pays $50 on signing.");
  });

  test("builds arguments from flags and --input, flags winning", async () => {
    const file = await writeDocx(dir, "contract.docx", CONTRACT_PARAGRAPHS);
    const captured = captureIo({ stdin: '{"query":"terminate","matchCase":true}' });

    const exit = await runFolioCli(["find", file, "--input", "-", "--query", "Late"], captured.io);

    expect(exit).toBe(0);
    const data = envelopeOf(captured.stdout())["data"];
    expect(JSON.stringify(data)).toContain('"totalMatches":1');
    expect(JSON.stringify(data)).toContain("Late payment");
  });

  test("maps failures to the envelope and exit classes", async () => {
    const file = await writeDocx(dir, "contract.docx", CONTRACT_PARAGRAPHS);
    const cases: [string[], number, string][] = [
      [["publish", file], 2, "usage_error"],
      [["read", file, "--no-such-flag"], 2, "usage_error"],
      [["read"], 2, "usage_error"],
      [["read", `${dir}/missing.docx`], 6, "not_found"],
      [["read", file, "--expect-version", "0".repeat(64)], 10, "stale_version"],
      [["read", file, "--max-blocks", "zero"], 2, "usage_error"],
    ];
    for (const [argv, exitCode, code] of cases) {
      const captured = captureIo();
      expect(await runFolioCli(argv, captured.io)).toBe(exitCode);
      const envelope = envelopeOf(captured.stdout());
      expect(envelope["ok"]).toBe(false);
      expect(JSON.stringify(envelope["error"])).toContain(`"code":"${code}"`);
    }
  });

  test("writes text-mode failures to stderr with the hint", async () => {
    const file = await writeDocx(dir, "contract.docx", CONTRACT_PARAGRAPHS);
    const captured = captureIo({ isTTY: true });

    const exit = await runFolioCli(["read", file, "--expect-version", "0".repeat(64)], captured.io);

    expect(exit).toBe(10);
    expect(captured.stdout()).toBe("");
    expect(captured.stderr()).toContain("error: ");
    expect(captured.stderr()).toContain("hint: Re-read the document");
  });

  test("prints root and command help", async () => {
    const root = captureIo();
    const command = captureIo();

    expect(await runFolioCli(["--help"], root.io)).toBe(0);
    expect(await runFolioCli(["find", "--help"], command.io)).toBe(0);

    expect(root.stdout()).toContain("Exit codes:");
    expect(root.stdout()).toContain("outline");
    expect(command.stdout()).toContain("--match-case");
    expect(command.stdout()).toContain("--expect-version");
  });
});
