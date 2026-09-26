import { afterAll, describe, expect, test } from "bun:test";
import { mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";

import {
  fileVersionOf,
  parseSaveEnvelope,
  saveArgs,
  saveWithCli,
  type CliSaveRequest,
} from "./save";

const VERSION = "a".repeat(64);

const request = (overrides: Partial<CliSaveRequest> = {}): CliSaveRequest => ({
  documentPath: "/work/Report.docx",
  bytes: new TextEncoder().encode("PK saved"),
  expectedVersion: VERSION,
  author: "Ada Lovelace",
  strategy: { type: "selective-first" },
  ...overrides,
});

describe("saveArgs", () => {
  test("saves in place against the baseline, as the VS Code editor", () => {
    expect(saveArgs(request(), "/tmp/saved.docx")).toEqual([
      "save",
      "/work/Report.docx",
      "--from",
      "/tmp/saved.docx",
      "--expect-version",
      VERSION,
      "--author",
      "Ada Lovelace",
      "--owner",
      "folio-vscode",
      "--surface",
      "vscode",
      "--save-strategy",
      "selective",
      "--output",
      "json",
    ]);
  });

  test("names the lease, the repack, and a destination to replace", () => {
    const args = saveArgs(
      request({
        leaseToken: "tok",
        strategy: { type: "full-repack", reason: "noBodyView" },
        destination: { path: "/work/Copy.docx", expectedVersion: "b".repeat(64) },
      }),
      "/tmp/saved.docx",
    );

    expect(args.slice(12)).toEqual([
      "--save-strategy",
      "full-repack",
      "--lease-token",
      "tok",
      "-o",
      "/work/Copy.docx",
      "--overwrite",
      "--expect-destination-version",
      "b".repeat(64),
      "--output",
      "json",
    ]);
  });

  test("writes a new destination without --overwrite", () => {
    const args = saveArgs(
      request({ destination: { path: "/work/New.docx", expectedVersion: null } }),
      "/tmp/saved.docx",
    );

    expect(args.slice(14, -2)).toEqual(["-o", "/work/New.docx"]);
  });
});

describe("parseSaveEnvelope", () => {
  test("reads the receipt's version, status, and backup", () => {
    const stdout = JSON.stringify({
      ok: true,
      data: { fileVersion: VERSION, status: "committed", backup: "/work/.folio/backups/x" },
    });

    expect(parseSaveEnvelope(stdout)).toEqual({
      type: "saved",
      fileVersion: VERSION,
      status: "committed",
      backup: "/work/.folio/backups/x",
    });
  });

  test("reads the error code, message, and hint", () => {
    const stdout = JSON.stringify({
      ok: false,
      error: { code: "stale_version", message: "changed", hint: "Reload." },
    });

    expect(parseSaveEnvelope(stdout)).toEqual({
      type: "error",
      code: "stale_version",
      message: "changed",
      hint: "Reload.",
    });
  });

  test("is null for anything else", () => {
    expect(parseSaveEnvelope("")).toBeNull();
    expect(parseSaveEnvelope("oops")).toBeNull();
    expect(parseSaveEnvelope(JSON.stringify({ ok: true, data: { status: "x" } }))).toBeNull();
    expect(parseSaveEnvelope(JSON.stringify({ ok: false, error: { message: "m" } }))).toBeNull();
  });
});

describe("saveWithCli", () => {
  const scripts = mkdtempSync(path.join(tmpdir(), "folio-save-test-"));
  afterAll(() => rmSync(scripts, { recursive: true, force: true }));

  const fakeCli = (name: string, body: string): string => {
    const file = path.join(scripts, `${name}.mjs`);
    writeFileSync(
      file,
      [
        'import fs from "node:fs";',
        "const args = process.argv.slice(2);",
        'const from = args[args.indexOf("--from") + 1];',
        body,
      ].join("\n"),
    );
    return file;
  };

  const save = (cliEntry: string, timeoutMs?: number) =>
    saveWithCli({ nodePath: process.execPath, cliEntry }, request(), timeoutMs);

  test("hands the CLI the bytes and returns its receipt", async () => {
    const cli = fakeCli(
      "ok",
      [
        'if (fs.readFileSync(from, "utf8") !== "PK saved") process.exit(65);',
        'if (args[0] !== "save" || args[1] !== "/work/Report.docx") process.exit(66);',
        "console.log(JSON.stringify({ ok: true, data: { fileVersion: 'f'.repeat(64), status: 'committed' } }));",
      ].join("\n"),
    );

    expect(await save(cli)).toEqual({
      type: "saved",
      fileVersion: "f".repeat(64),
      status: "committed",
    });
  });

  test("returns the CLI's refusal", async () => {
    const cli = fakeCli(
      "stale",
      [
        "console.log(JSON.stringify({ ok: false, error: { code: 'stale_version', message: 'changed' } }));",
        "process.exit(5);",
      ].join("\n"),
    );

    expect(await save(cli)).toEqual({ type: "error", code: "stale_version", message: "changed" });
  });

  test("reports a crash with the end of its stderr", async () => {
    const cli = fakeCli("crash", 'console.error("boom");\nprocess.exit(9);');

    expect(await save(cli)).toEqual({
      type: "error",
      code: "crashed",
      message: "folio save exited with code 9.\nboom",
    });
  });

  test("gives up after the timeout", async () => {
    const cli = fakeCli("hang", "setTimeout(() => {}, 10_000);");

    const outcome = await save(cli, 100);

    expect(outcome.type).toBe("error");
    if (outcome.type === "error") expect(outcome.code).toBe("crashed");
  });
});

describe("fileVersionOf", () => {
  test("is the SHA-256 of the bytes in hex, as folio's", () => {
    expect(fileVersionOf(new TextEncoder().encode("abc"))).toBe(
      "ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad",
    );
  });
});
