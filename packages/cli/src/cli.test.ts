import { afterEach, beforeEach, describe, expect, test } from "bun:test";

import {
  CONTRACT_PARAGRAPHS,
  makeTempDir,
  writeDocx,
} from "./__tests__/fixtures";
import { captureIo, dataOf, envelopeOf } from "./__tests__/io";
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
  test("deletes comments by repeated id and exact author", async () => {
    const file = await writeDocx(dir, "comments.docx", CONTRACT_PARAGRAPHS);
    const run = async (command: string, ...flags: string[]) => {
      const captured = captureIo();
      expect(
        await runFolioCli([command, file, ...flags], captured.io),
        captured.stdout(),
      ).toBe(0);
      return dataOf(captured.stdout());
    };
    const writeFlags = ["--in-place", "--no-expect-version", "--allow-repack"];
    const add = async (author: string, text: string) => {
      await run(
        "comment",
        "--author",
        author,
        "--block-id",
        "10000002",
        "--text",
        text,
        ...writeFlags,
      );
    };

    await add("Reviewer", "First");
    await add("Reviewer", "Second");
    await add("Reviewer", "Third");
    await add("Reviewer Jr.", "Keep");
    const comments = (await run("comments"))["result"];
    if (!Array.isArray(comments)) throw new Error("Missing comments");
    const idOf = (text: string) => {
      const comment: unknown = comments.find(
        (entry: unknown) =>
          typeof entry === "object" &&
          entry !== null &&
          "text" in entry &&
          entry.text === text,
      );
      if (
        typeof comment !== "object" ||
        comment === null ||
        !("id" in comment)
      ) {
        throw new Error(`Missing comment ${text}`);
      }
      return String(comment.id);
    };
    const first = idOf("First");
    const second = idOf("Second");
    const selectedByAuthor = idOf("Third");

    const byId = await run(
      "delete-comments",
      "--id",
      first,
      "--id",
      second,
      ...writeFlags,
    );
    expect(byId["result"]).toEqual({
      deleted: [Number(first), Number(second)],
    });

    const byAuthor = await run(
      "delete-comments",
      "--by-author",
      "Reviewer",
      ...writeFlags,
    );
    expect(byAuthor["result"]).toEqual({ deleted: [Number(selectedByAuthor)] });
    expect((await run("comments"))["result"]).toMatchObject([
      { author: "Reviewer Jr.", text: "Keep" },
    ]);
  });

  test.each([
    { selector: [] },
    { selector: ["--id", "1", "--by-author", "Reviewer"] },
    { selector: ["--id", "1.5"] },
    { selector: ["--id", "9007199254740992"] },
  ])(
    "refuses an invalid comment selection $selector without changing the file",
    async ({ selector }) => {
      const file = await writeDocx(dir, "unchanged.docx", CONTRACT_PARAGRAPHS);
      const before = await Bun.file(file).arrayBuffer();
      const captured = captureIo();
      const exit = await runFolioCli(
        [
          "delete-comments",
          file,
          "--in-place",
          "--no-expect-version",
          ...selector,
        ],
        captured.io,
      );
      expect(exit).not.toBe(0);
      expect(envelopeOf(captured.stdout())["error"]).toMatchObject({
        code: "invalid_input",
      });
      expect(await Bun.file(file).arrayBuffer()).toEqual(before);
    },
  );

  test("prints a JSON envelope off a terminal and text on one", async () => {
    const file = await writeDocx(dir, "contract.docx", CONTRACT_PARAGRAPHS);
    const piped = captureIo();
    const terminal = captureIo({ isTTY: true });

    const pipedExit = await runFolioCli(["read", file], piped.io);
    const terminalExit = await runFolioCli(["read", file], terminal.io);

    expect([pipedExit, terminalExit]).toEqual([0, 0]);
    expect(envelopeOf(piped.stdout())["ok"]).toBe(true);
    expect(terminal.stdout()).toContain(
      "[10000002] The buyer pays $50 on signing.",
    );
  });

  test("builds arguments from flags and --input, flags winning", async () => {
    const file = await writeDocx(dir, "contract.docx", CONTRACT_PARAGRAPHS);
    const captured = captureIo({
      stdin: '{"query":"terminate","matchCase":true}',
    });

    const exit = await runFolioCli(
      ["find", file, "--input", "-", "--query", "Late"],
      captured.io,
    );

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

    const exit = await runFolioCli(
      ["read", file, "--expect-version", "0".repeat(64)],
      captured.io,
    );

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

  test("shows a worked operation for suggest", async () => {
    const help = captureIo();

    expect(await runFolioCli(["suggest", "--help"], help.io)).toBe(0);

    expect(help.stdout()).toContain('"type":"replaceInBlock"');
    expect(help.stdout()).toContain("folio find contract.docx --query");
    expect(help.stdout()).toContain("--existing-revisions");
    expect(help.stdout()).toContain("refuse-cross-author");
  });

  test("refuses an edit in another author's revision when requested", async () => {
    const file = await writeDocx(dir, "contract.docx", CONTRACT_PARAGRAPHS);
    const suggest = ({
      find,
      replacement,
      author,
      policy = [],
    }: {
      find: string;
      replacement: string;
      author: string;
      policy?: readonly string[];
    }) => {
      const captured = captureIo();
      const exit = runFolioCli(
        [
          "suggest",
          file,
          "--in-place",
          "--no-expect-version",
          "--author",
          author,
          "--operations",
          JSON.stringify([
            {
              type: "replaceInBlock",
              blockId: "10000002",
              find,
              replace: replacement,
            },
          ]),
          ...policy,
        ],
        captured.io,
      );
      return { captured, exit };
    };
    const first = suggest({
      find: "pays",
      replacement: "tenders",
      author: "Reviewer One",
    });
    expect(await first.exit, first.captured.stdout()).toBe(0);
    const before = await Bun.file(file).arrayBuffer();
    const second = suggest({
      find: "tenders",
      replacement: "delivers",
      author: "Reviewer Two",
      policy: ["--existing-revisions", "refuse-cross-author"],
    });
    expect(await second.exit).not.toBe(0);
    expect(second.captured.stdout()).toContain("existingRevisionConflict");
    expect(await Bun.file(file).arrayBuffer()).toEqual(before);
  });
});
