import { afterAll, describe, expect, test } from "bun:test";
import { mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";

import { parseRenderEnvelope, renderDocument, stderrTail } from "./render";

describe("parseRenderEnvelope", () => {
  test("reads the page count from a success envelope", () => {
    const stdout = `${JSON.stringify({ ok: true, data: { format: "html", pageCount: 4 } })}\n`;

    expect(parseRenderEnvelope(stdout)).toEqual({ type: "ok", pageCount: 4 });
  });

  test("reads the message and hint from a failure envelope", () => {
    const stdout = JSON.stringify({
      ok: false,
      error: { code: "invalid_document", message: "not a package", hint: "open it and save" },
    });

    expect(parseRenderEnvelope(stdout)).toEqual({
      type: "error",
      message: "not a package",
      hint: "open it and save",
    });
    expect(parseRenderEnvelope(JSON.stringify({ ok: false, error: { message: "m" } }))).toEqual({
      type: "error",
      message: "m",
    });
  });

  test("reads the last line when something printed before it", () => {
    const stdout = `warming up\n${JSON.stringify({ ok: true, data: { pageCount: 1 } })}\n\n`;

    expect(parseRenderEnvelope(stdout)).toEqual({ type: "ok", pageCount: 1 });
  });

  test("is null for anything that is not an envelope", () => {
    expect(parseRenderEnvelope("")).toBeNull();
    expect(parseRenderEnvelope("Segmentation fault")).toBeNull();
    expect(parseRenderEnvelope("[1,2]")).toBeNull();
    expect(parseRenderEnvelope(JSON.stringify({ ok: true, data: {} }))).toBeNull();
    expect(parseRenderEnvelope(JSON.stringify({ ok: true, data: { pageCount: "3" } }))).toBeNull();
    expect(parseRenderEnvelope(JSON.stringify({ ok: false, error: {} }))).toBeNull();
  });
});

describe("stderrTail", () => {
  test("keeps the last lines", () => {
    expect(stderrTail("a\nb\nc\nd\n", 2)).toBe("c\nd");
  });
});

describe("renderDocument", () => {
  const scripts = mkdtempSync(path.join(tmpdir(), "folio-render-test-"));
  afterAll(() => rmSync(scripts, { recursive: true, force: true }));

  /** A stand-in CLI: `body` sees `input`, `output`, and `fs`. */
  const fakeCli = (name: string, body: string): string => {
    const file = path.join(scripts, `${name}.mjs`);
    writeFileSync(
      file,
      [
        'import fs from "node:fs";',
        "const [command, input, , output] = process.argv.slice(2);",
        'if (command !== "render") process.exit(64);',
        body,
      ].join("\n"),
    );
    return file;
  };

  const render = (cliEntry: string, signal = new AbortController().signal, timeoutMs?: number) =>
    renderDocument({
      runtime: { nodePath: process.execPath, cliEntry },
      bytes: new TextEncoder().encode("PK fake"),
      fileName: "Report.docx",
      signal,
      ...(timeoutMs !== undefined && { timeoutMs }),
    });

  test("returns the HTML the CLI wrote and its page count", async () => {
    const cli = fakeCli(
      "ok",
      [
        'if (fs.readFileSync(input, "utf8") !== "PK fake") process.exit(65);',
        'if (process.env.ELECTRON_RUN_AS_NODE !== "1") process.exit(66);',
        'fs.writeFileSync(output, "<html>pages</html>");',
        "console.log(JSON.stringify({ ok: true, data: { pageCount: 2 } }));",
      ].join("\n"),
    );

    expect(await render(cli)).toEqual({
      type: "document",
      html: "<html>pages</html>",
      pageCount: 2,
    });
  });

  test("reports the CLI's error under the document's own name", async () => {
    const cli = fakeCli(
      "refused",
      [
        "console.log(JSON.stringify({ ok: false, error: {",
        "  code: 'invalid_document', message: `${input} is not a .docx package.`, hint: 'Re-save it.' } }));",
        "process.exit(3);",
      ].join("\n"),
    );

    expect(await render(cli)).toEqual({
      type: "error",
      message: "Report.docx is not a .docx package.",
      hint: "Re-save it.",
    });
  });

  test("reports a crash with the end of its stderr", async () => {
    const cli = fakeCli("crash", 'console.error("boom");\nprocess.exit(9);');

    expect(await render(cli)).toEqual({
      type: "error",
      message: "The renderer exited with code 9.\nboom",
    });
  });

  test("is cancelled when the signal aborts", async () => {
    const cli = fakeCli("slow", "setTimeout(() => {}, 10_000);");
    const controller = new AbortController();

    const pending = render(cli, controller.signal);
    setTimeout(() => controller.abort(), 50);

    expect(await pending).toEqual({ type: "cancelled" });
  });

  test("gives up after the timeout", async () => {
    const cli = fakeCli("hang", "setTimeout(() => {}, 10_000);");

    const outcome = await render(cli, new AbortController().signal, 100);

    expect(outcome.type).toBe("error");
  });
});
