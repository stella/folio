/**
 * Every command that takes a document refuses a file that is not a
 * WordprocessingML package, before it reads a block or writes a byte.
 *
 * A ZIP of unrelated entries used to read as an empty document and commit as
 * a successful `save`, replacing a real package; the invalid-input tests only
 * used bytes that are not a ZIP at all, which JSZip already refuses. The
 * matrix is every malformed shape crossed with every command, and the command
 * list comes from the registry and the help text, so a new command is covered
 * the day it is added.
 */

import { afterAll, beforeAll, describe, expect, test } from "bun:test";
import { readFile, writeFile } from "node:fs/promises";
import path from "node:path";

import { buildDocx, CONTRACT_PARAGRAPHS, makeTempDir } from "./__tests__/fixtures";
import { captureIo, envelopeOf } from "./__tests__/io";
import { runFolioCli } from "./cli";
import { fileVersionOf } from "./document";
import {
  MALFORMED_PACKAGES,
  TOOL_ARGUMENTS,
  WELL_FORMED_PACKAGES,
} from "./__tests__/malformed-packages";
import { checkWordprocessingPackage, INVALID_PACKAGE_REASONS } from "./main-document-part";
import { listCommands } from "./registry";

let dir = "";
let cleanup: () => Promise<void> = () => Promise.resolve();
let valid: Uint8Array = new Uint8Array();

beforeAll(async () => {
  ({ dir, cleanup } = await makeTempDir());
  valid = await buildDocx(CONTRACT_PARAGRAPHS);
});

afterAll(async () => {
  await cleanup();
});

describe("checkWordprocessingPackage", () => {
  test("names the reason for every malformed shape", async () => {
    for (const { name, reason, build } of MALFORMED_PACKAGES) {
      const checked = await checkWordprocessingPackage(name, await build(valid));
      expect(checked.isErr(), name).toBe(true);
      if (checked.isOk()) continue;
      expect(checked.error.code, name).toBe("invalid_document");
      expect(checked.error.details, name).toEqual({ reason });
    }
  });

  test("the malformed matrix exercises every reason", () => {
    expect(new Set(MALFORMED_PACKAGES.map(({ reason }) => reason))).toEqual(
      new Set(Object.values(INVALID_PACKAGE_REASONS)),
    );
  });

  test.each(WELL_FORMED_PACKAGES)("accepts $name", async ({ name, build }) => {
    const checked = await checkWordprocessingPackage(name, await build(valid));
    expect(checked.isOk()).toBe(true);
  });
});

/** File commands `runFolioCli` dispatches outside the registry, and how to aim them at a file. */
const NON_REGISTRY_COMMANDS: Readonly<
  Record<string, (file: { path: string; version: string }, other: string) => string[]>
> = {
  save: ({ path: filePath, version }, other) => [
    "save",
    other,
    "--from",
    filePath,
    "--expect-version",
    version,
  ],
  render: ({ path: filePath }) => ["render", filePath, "-o", `${filePath}.html`],
  serve: ({ path: filePath }) => ["serve", filePath],
};

/** Commands `folio --help` lists that never read a document. */
const NO_FILE_COMMANDS = new Set(["mcp"]);

const helpCommands = async (): Promise<string[]> => {
  const captured = captureIo();
  await runFolioCli(["--help"], captured.io);
  return [...captured.stdout().matchAll(/^ {2}([a-z]+) {2,}/gmu)].map(([, name]) => name ?? "");
};

type Invocation = { label: string; argv: string[]; stdin?: string };

const invocationsFor = (malformed: string, other: string): Invocation[] => {
  const version = fileVersionOf(new Uint8Array());
  const invocations: Invocation[] = [];
  for (const { tool, command } of listCommands()) {
    if (tool.type === "compare") {
      invocations.push(
        { label: `${command.name} (base)`, argv: [command.name, malformed, other] },
        { label: `${command.name} (revised)`, argv: [command.name, other, malformed] },
        {
          label: `${command.name} (redline)`,
          argv: [
            command.name,
            malformed,
            other,
            "-o",
            `${malformed}.redline.docx`,
            "--no-expect-version",
          ],
        },
      );
      continue;
    }
    const args = TOOL_ARGUMENTS[tool.name];
    if (args === undefined) throw new Error(`no arguments for ${tool.name}`);
    const write = tool.type !== "agentRead";
    invocations.push({
      label: command.name,
      argv: [
        command.name,
        malformed,
        "--input",
        "-",
        ...(write ? ["--no-expect-version", "--in-place"] : []),
      ],
      stdin: JSON.stringify(args),
    });
  }
  for (const [name, argvFor] of Object.entries(NON_REGISTRY_COMMANDS)) {
    invocations.push({ label: name, argv: argvFor({ path: malformed, version }, other) });
  }
  const saveOnto = NON_REGISTRY_COMMANDS["save"];
  if (saveOnto !== undefined) {
    // The other direction: a valid package saved over a malformed file.
    invocations.push({
      label: "save (onto)",
      argv: ["save", malformed, "--from", other, "--expect-version", version],
    });
  }
  return invocations;
};

describe("every command refuses a malformed package", () => {
  test("the matrix covers every tool and every command the help lists", async () => {
    const toolNames = new Set(listCommands().map(({ tool }) => tool.name));
    expect(new Set(Object.keys(TOOL_ARGUMENTS))).toEqual(
      new Set([...toolNames].filter((name) => name !== "compare_documents")),
    );
    const fileCommands = (await helpCommands()).filter((name) => !NO_FILE_COMMANDS.has(name));
    expect(new Set(fileCommands)).toEqual(
      new Set([
        ...listCommands().map(({ command }) => command.name),
        ...Object.keys(NON_REGISTRY_COMMANDS),
      ]),
    );
  });

  test.each(MALFORMED_PACKAGES)(
    "$name",
    async ({ name, build }) => {
      const slug = name.replaceAll(/[^a-z]+/gu, "-");
      const malformed = path.join(dir, `${slug}.docx`);
      const other = path.join(dir, `${slug}-valid.docx`);
      const bytes = await build(valid);
      await writeFile(malformed, bytes);
      await writeFile(other, valid);
      const outcomes: string[] = [];
      for (const { label, argv, stdin } of invocationsFor(malformed, other)) {
        const captured = captureIo({ ...(stdin !== undefined && { stdin }), cwd: dir });
        // `serve` runs until interrupted; one that accepted the file stops at once.
        const io = { ...captured.io, untilInterrupted: () => Promise.resolve() };
        const exit = await runFolioCli(argv, io);
        const error = envelopeOf(captured.stdout())["error"];
        const code =
          typeof error === "object" && error !== null ? Reflect.get(error, "code") : null;
        outcomes.push(`${label}: ${String(exit)} ${String(code)}`);
      }
      expect(outcomes).toEqual(
        invocationsFor(malformed, other).map(({ label }) => `${label}: 2 invalid_document`),
      );
      // Nothing was written over either file.
      expect(new Uint8Array(await readFile(malformed))).toEqual(bytes);
      expect(new Uint8Array(await readFile(other))).toEqual(valid);
    },
    60_000,
  );
});
