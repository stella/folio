import { expect, test } from "bun:test";
import { mkdtemp, rm } from "node:fs/promises";
import path from "node:path";
import { tmpdir } from "node:os";

const pluginPath = path.resolve(import.meta.dir, "oxlint.ts");
const oxlintPath = path.resolve(import.meta.dir, "../../../node_modules/.bin/oxlint");

const lint = async (
  rule: string,
  source: string,
  { filename = "source.ts", options }: { filename?: string; options?: unknown } = {},
) => {
  const directory = await mkdtemp(path.join(tmpdir(), "utf8-boundaries-"));
  try {
    const configPath = path.join(directory, "oxlint.json");
    const sourcePath = path.join(directory, filename);
    await Bun.write(
      configPath,
      JSON.stringify({
        categories: { correctness: "off" },
        jsPlugins: [pluginPath],
        rules: {
          [`utf8-boundaries/${rule}`]: options === undefined ? "error" : ["error", options],
        },
      }),
    );
    await Bun.write(sourcePath, source);
    const child = Bun.spawn(
      [process.execPath, "--bun", oxlintPath, "-c", configPath, "-f", "json", sourcePath],
      {
        cwd: directory,
        stdout: "pipe",
        stderr: "pipe",
      },
    );
    const [stdout, stderr, exitCode] = await Promise.all([
      new Response(child.stdout).text(),
      new Response(child.stderr).text(),
      child.exited,
    ]);
    if (exitCode !== 0 && exitCode !== 1) throw new Error(`${stdout}\n${stderr}`);
    const report: unknown = JSON.parse(stdout);
    if (
      typeof report !== "object" ||
      report === null ||
      !("diagnostics" in report) ||
      !Array.isArray(report.diagnostics)
    ) {
      throw new Error(`Unexpected oxlint report: ${stdout}`);
    }
    const ruleCode = `utf8-boundaries(${rule})`;
    const relevantDiagnostics = report.diagnostics.filter(
      (diagnostic: unknown) =>
        typeof diagnostic === "object" &&
        diagnostic !== null &&
        "code" in diagnostic &&
        typeof diagnostic.code === "string" &&
        diagnostic.code.startsWith(ruleCode),
    );
    if (relevantDiagnostics.length !== report.diagnostics.length) {
      throw new Error(`Unexpected diagnostics outside ${ruleCode}: ${stdout}`);
    }
    return relevantDiagnostics
      .map((diagnostic: unknown) => {
        if (
          typeof diagnostic !== "object" ||
          diagnostic === null ||
          !("labels" in diagnostic) ||
          !Array.isArray(diagnostic.labels)
        ) {
          throw new Error(`Unexpected diagnostic: ${JSON.stringify(diagnostic)}`);
        }
        const label = diagnostic.labels[0];
        if (
          typeof label !== "object" ||
          label === null ||
          !("span" in label) ||
          typeof label.span !== "object" ||
          label.span === null ||
          !("line" in label.span) ||
          typeof label.span.line !== "number"
        ) {
          throw new Error(`Unexpected diagnostic span: ${JSON.stringify(diagnostic)}`);
        }
        return label.span.line;
      })
      .toSorted((left: number, right: number) => left - right);
  } finally {
    await rm(directory, { recursive: true, force: true });
  }
};

test("no-string-zip-entry catches computed writes and permits Bun and literal callbacks", async () => {
  const source = [
    'zip.file("doc.xml", xml);',
    'zip["file"](path, content);',
    "Bun.file(path, options);",
    "entry.file((file) => accept(file), (error) => reject(error));",
    "entry.file(function (file) {}, onError);",
    "zip.file(path);",
  ].join("\n");
  expect(await lint("no-string-zip-entry", source)).toEqual([1, 2]);
});

test("no-string-zip-entry exempts only the configured owner path", async () => {
  const source = "zip.file(path, content);";
  const options = { ownerPath: "src/zip-owner.ts" };
  expect(
    await lint("no-string-zip-entry", source, { filename: "src/zip-owner.ts", options }),
  ).toEqual([]);
  expect(
    await lint("no-string-zip-entry", source, { filename: "src/consumer.ts", options }),
  ).toEqual([1]);
  expect(
    await lint("no-string-zip-entry", source, { filename: "other-src/zip-owner.ts", options }),
  ).toEqual([1]);
});

test("no-chunked-text-codec tracks decoder receivers across loops and stream callbacks", async () => {
  const source = [
    "const decoder = new TextDecoder();",
    "for (const chunk of chunks) decoder.decode(chunk);",
    "for (const chunk of chunks) decoder.decode(chunk, { stream: true });",
    "for (const chunk of chunks) new TextDecoder().decode(chunk);",
    "for (const chunk of chunks) for (const part of chunk) decoder.decode(part);",
    'stream.on("data", (chunk) => decoder.decode(chunk));',
    'stream.on("data", (chunk) => decoder.decode(chunk, { stream: true }));',
    "decoder.decode(finalChunk);",
    "codec.decode(chunk);",
  ].join("\n");
  const withObjectHandlers = [
    "const decoder = new TextDecoder();",
    "new TransformStream({ transform(chunk) { decoder.decode(chunk); } });",
    "new ReadableStream({ start(controller) { decoder.decode(chunk); }, pull(controller) { decoder.decode(chunk); } });",
    "decoder.decode(finalChunk);",
  ].join("\n");
  expect(await lint("no-chunked-text-codec", source)).toEqual([2, 4, 4, 5, 6]);
  expect(await lint("no-chunked-text-codec", withObjectHandlers)).toEqual([2, 3, 3]);
});

test("fixtures exercise both rules", async () => {
  const zipFixture = await Bun.file(
    path.resolve(import.meta.dir, "../test/fixtures/no-string-zip-entry.ts"),
  ).text();
  const codecFixture = await Bun.file(
    path.resolve(import.meta.dir, "../test/fixtures/no-chunked-text-codec.ts"),
  ).text();
  expect(await lint("no-string-zip-entry", zipFixture)).toEqual([2, 3, 4, 5, 6, 7]);
  expect(await lint("no-chunked-text-codec", codecFixture)).toEqual([3, 4, 5, 5, 6]);
});

test("stream options must remain enabled after spreads and computed keys", async () => {
  const source = [
    "const decoder = new TextDecoder();",
    "for (const chunk of chunks) decoder.decode(chunk, { stream: true, ...options });",
    "for (const chunk of chunks) decoder.decode(chunk, { ...options, stream: true });",
    "for (const chunk of chunks) decoder.decode(chunk, { stream: true, [key]: false });",
    "for (const chunk of chunks) decoder.decode(chunk, { stream: true, fatal: true });",
    "for (const chunk of chunks) decoder.decode(chunk, { ['stream']: true });",
  ].join("\n");
  expect(await lint("no-chunked-text-codec", source)).toEqual([2, 4]);
});
