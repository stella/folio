/** One interleaved run, used unchanged by the quiet window and draft-PR CI. */
import { chromium } from "@playwright/test";
import { loadavg, cpus, platform, arch, totalmem } from "node:os";
import { spawnSync } from "node:child_process";
import { mkdirSync, appendFileSync, readFileSync, writeFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { resolve } from "node:path";
import { isDeepStrictEqual } from "node:util";
import { gzipSync } from "node:zlib";
import { createHash } from "node:crypto";
import { projectionPatch } from "./resident-projection";
import { sourceDigest } from "./artifacts";
import { applyDocumentOps } from "../../../packages/docx-core/src/ops/apply";
import type { Document, Paragraph, TableRow } from "../../../packages/docx-core/src/model/document";
import { DOCUMENT_OP_TYPES, type DocumentOp } from "../../../packages/docx-core/src/ops/types";
import {
  allocateEditorIntentIds,
  compileEditorIntent,
} from "../../../packages/docx-core/src/ops/editorIntent";

const root = fileURLToPath(new URL("..", import.meta.url));
const progress = process.env["RUST_SPIKE_PROGRESS_DIR"];
if (!progress) throw new TypeError("RUST_SPIKE_PROGRESS_DIR is required.");
mkdirSync(progress, { recursive: true });
const mode = process.env["RUST_SPIKE_BENCH_MODE"];
if (!["quiet-window", "ci", "validate"].includes(mode ?? ""))
  throw new TypeError("Explicit benchmark mode is required.");
const quiet = mode === "quiet-window";
const validateOnly = mode === "validate";
if (mode === "ci" && !process.env["CI"])
  throw new TypeError("CI measurements require a CI runner.");
if (quiet && (loadavg().at(0) ?? Infinity) >= 3)
  throw new TypeError("Quiet-window start load is not below 3; exiting without retry.");
const output = resolve(
  progress,
  validateOnly ? "fixture-validation.jsonl" : `${quiet ? "quiet" : "ci"}-measurements.jsonl`,
);
// A run owns its evidence file; existing samples must never be appended or overwritten.
writeFileSync(output, "", { flag: validateOnly ? "w" : "wx" });
const nativeBinary = resolve(root, "target/release/canonical-spike");
const manifest: unknown = JSON.parse(
  readFileSync(
    resolve(progress, quiet ? "benchmark-ready.json" : "benchmark-artifacts.json"),
    "utf8",
  ),
);
if (
  typeof manifest !== "object" ||
  manifest === null ||
  !("status" in manifest) ||
  manifest.status !== (quiet ? "validated" : "built") ||
  !("sourceHash" in manifest) ||
  manifest.sourceHash !== sourceDigest(root) ||
  !("artifacts" in manifest) ||
  !Array.isArray(manifest.artifacts)
)
  throw new TypeError("Artifacts do not match current source.");
for (const artifact of manifest.artifacts) {
  if (
    typeof artifact?.path !== "string" ||
    typeof artifact.sha256 !== "string" ||
    createHash("sha256").update(readFileSync(artifact.path)).digest("hex") !== artifact.sha256
  )
    throw new TypeError("Artifact hash mismatch.");
}
const assets = new Map([
  [
    "/",
    new Response('<script type="module" src="/browser.js"></script>', {
      headers: { "Content-Type": "text/html" },
    }),
  ],
  [
    "/browser.js",
    new Response(readFileSync(resolve(root, "target/benchmark-browser/browser.js")), {
      headers: { "Content-Type": "text/javascript" },
    }),
  ],
  [
    "/wasm/docx_canonical_spike.js",
    new Response(
      readFileSync(resolve(root, "target/wasm-bindgen-release/docx_canonical_spike.js")),
      { headers: { "Content-Type": "text/javascript" } },
    ),
  ],
  [
    "/wasm/docx_canonical_spike_bg.wasm",
    new Response(
      readFileSync(resolve(root, "target/wasm-bindgen-release/docx_canonical_spike_bg.wasm")),
      { headers: { "Content-Type": "application/wasm" } },
    ),
  ],
]);
const record = (value: unknown) => appendFileSync(output, `${JSON.stringify(value)}\n`);
record({
  type: "runner",
  time: `${new Date().toLocaleString("sv-SE", { timeZone: "Europe/Prague" })} CEST`,
  platform: platform(),
  arch: arch(),
  cpu: cpus().at(0)?.model,
  cores: cpus().length,
  memoryBytes: totalmem(),
  buildRevision: Reflect.get(manifest, "revision"),
  sourceHash: Reflect.get(manifest, "sourceHash"),
  runnerName: process.env["RUNNER_NAME"],
  load1: loadavg().at(0),
  revision:
    process.env["GITHUB_SHA"] ??
    spawnSync("git", ["rev-parse", "HEAD"], { encoding: "utf8" }).stdout.trim(),
  mode,
});

const fixturesForPages = (pages: number) => {
  const paragraph = (index: number) =>
    ({
      type: "paragraph",
      paraId: (index + 1).toString(16).toUpperCase().padStart(8, "0"),
      ...(index > 0 ? { formatting: { pageBreakBefore: true } } : {}),
      content: [
        {
          type: "run",
          content: [
            {
              type: "text",
              text: "A synthetic paragraph with enough text to model editing across one hundred pages. ".repeat(
                8,
              ),
            },
          ],
        },
      ],
    }) satisfies Paragraph;
  // Explicit page breaks model page count; this does not benchmark layout/pagination.
  const document = {
    package: {
      document: { content: Array.from({ length: pages }, (_, index) => paragraph(index)) },
    },
  } satisfies Document;
  const blockId = document.package.document.content.at(0)?.paraId;
  if (!blockId) throw new TypeError("Benchmark document needs a leading paragraph.");
  const at = { story: "main", blockId, offset: 10 } as const;
  const to = { ...at, offset: 19 };
  const textOps = [
    { type: DOCUMENT_OP_TYPES.INSERT_TEXT, at, text: "benchmark", runProps: "inherit" },
    { type: DOCUMENT_OP_TYPES.DELETE_RANGE, from: at, to },
    { type: DOCUMENT_OP_TYPES.DELETE_BLOCKS, story: "main", blockIds: [blockId] },
    { type: DOCUMENT_OP_TYPES.SET_RUN_PROPS, from: at, to, patch: { bold: true } },
    {
      type: DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS,
      story: "main",
      blockId,
      patch: { alignment: "center" },
    },
    { type: DOCUMENT_OP_TYPES.SPLIT_BLOCK, at, newBlockId: "10000001" },
    {
      type: DOCUMENT_OP_TYPES.INSERT_BLOCKS,
      story: "main",
      at: { type: "after", blockId },
      blocks: [paragraph(0x10000001)],
    },
  ] as const satisfies readonly DocumentOp[];
  const operations = [
    ...textOps,
    ...(pages > 1
      ? [
          {
            type: DOCUMENT_OP_TYPES.JOIN_BLOCKS,
            story: "main",
            blockId,
            nextBlockId: "00000002",
          } as const,
        ]
      : []),
  ] satisfies readonly DocumentOp[];
  const cases = operations.map((op) => ({
    name: op.type,
    task: "apply" as const,
    document,
    ops: [op],
  }));
  const revision = { id: 100, author: "A", date: "2026-01-02T03:04:05Z" };
  const trackedCases = [
    { type: DOCUMENT_OP_TYPES.INSERT_TEXT, at, text: "benchmark", runProps: "inherit", revision },
    { type: DOCUMENT_OP_TYPES.DELETE_RANGE, from: at, to, revision },
  ] as const satisfies readonly DocumentOp[];
  const intent = { type: "replaceText", from: at, to, text: "benchmark" } as const;
  const allocation = allocateEditorIntentIds(document, intent);
  const replacement = compileEditorIntent(document, {
    intent,
    mode: { type: "editing", newIds: allocation.newIds },
  }).unwrap();
  const row = (index: number) =>
    ({
      type: "tableRow",
      cells: [{ type: "tableCell", content: [paragraph(index)] }],
    }) satisfies TableRow;
  const tableDocument = {
    package: {
      document: {
        content: [
          ...document.package.document.content,
          { type: "table", columnWidths: [2400], rows: [row(0x10000100), row(0x10000101)] },
          paragraph(0x10000200),
        ],
      },
    },
  } satisfies Document;
  const rowOps = [
    {
      type: DOCUMENT_OP_TYPES.INSERT_ROW,
      story: "main",
      blockId: "10000101",
      at: 1,
      row: row(0x10000102),
    },
    { type: DOCUMENT_OP_TYPES.DELETE_ROW, story: "main", blockId: "10000101" },
  ] as const satisfies readonly DocumentOp[];
  const extraCases = [
    ...trackedCases.map((op) => ({
      name: `tracked-${op.type}`,
      task: "apply" as const,
      document,
      ops: [op],
    })),
    { name: "replaceText-compiled-batch", task: "apply" as const, document, ops: replacement.ops },
    ...rowOps.map((op) => ({
      name: op.type,
      task: "apply" as const,
      document: tableDocument,
      ops: [op],
    })),
  ];
  return [
    ...[...cases, ...extraCases].map(({ name, task, document: caseDocument, ops }) => ({
      name,
      task,
      pages,
      document: caseDocument,
      ops,
      expected: applyDocumentOps(caseDocument, ops).unwrap(),
      documentJson: JSON.stringify(caseDocument),
      opsJson: JSON.stringify(ops),
    })),
    ...(["load", "save"] as const).map((task) => ({
      name: `json-model-${task}`,
      task,
      pages,
      document,
      ops: [],
      expected: null,
      documentJson: JSON.stringify(document),
      opsJson: "[]",
    })),
  ];
};
const fixtures = [1, 100].flatMap(fixturesForPages);
const wasmBytes = readFileSync(
  resolve(root, "target/wasm-bindgen-release/docx_canonical_spike_bg.wasm"),
);
record({
  type: "artifact-size",
  wasmRawBytes: wasmBytes.byteLength,
  wasmGzipBytes: gzipSync(wasmBytes).byteLength,
  nativeBytes: readFileSync(nativeBinary).byteLength,
});

const server = Bun.serve({
  port: 0,
  hostname: "127.0.0.1",
  fetch: (request) =>
    assets.get(new URL(request.url).pathname)?.clone() ??
    new Response("Not found", { status: 404 }),
});
let browser;
try {
  browser = await chromium.launch({ headless: true });
  const page = await browser.newPage();
  await page.goto(server.url.href);
  await page.waitForFunction(() => Reflect.get(globalThis, "canonicalSpikeReady") === true);
  record({ type: "browser", version: browser.version() });
  // Separate pages prevent retained resident models inflating whole-document WASM memory.
  const residentPage = await browser.newPage();
  await residentPage.goto(server.url.href);
  await residentPage.waitForFunction(() => Reflect.get(globalThis, "canonicalSpikeReady") === true);

  for (const fixture of fixtures) {
    // Verify native output before taking any timing; unsupported cases fail the run.
    if (fixture.task === "apply") {
      const undo = applyDocumentOps(fixture.expected.document, fixture.expected.inverse).unwrap();
      const redo = applyDocumentOps(undo.document, undo.inverse).unwrap();
      const checks = [
        { document: fixture.document, ops: fixture.ops, expected: fixture.expected },
        { document: fixture.expected.document, ops: fixture.expected.inverse, expected: undo },
        { document: undo.document, ops: undo.inverse, expected: redo },
      ];
      for (const check of checks) {
        const expected = JSON.parse(JSON.stringify(check.expected));
        const nativeCheck = spawnSync(nativeBinary, [], {
          input: `${JSON.stringify({ document: check.document, ops: check.ops })}\n`,
          encoding: "utf8",
        });
        if (nativeCheck.error) throw nativeCheck.error;
        if (
          nativeCheck.status !== 0 ||
          !isDeepStrictEqual(JSON.parse(nativeCheck.stdout), expected)
        )
          throw new TypeError(`Native fixture ${fixture.name} or inverse differs from TS.`);
        const expectedResident = {
          inverse: check.expected.inverse,
          touched: check.expected.touched,
          revisions: check.expected.revisions,
          projectionPatch: projectionPatch(check.document, check.expected.document),
        };
        const residentNative = spawnSync(nativeBinary, [], {
          input: `${JSON.stringify({ benchmark: { documentJson: JSON.stringify(check.document), opsJson: JSON.stringify(check.ops), task: "resident", verify: true } })}\n`,
          encoding: "utf8",
        });
        if (residentNative.error) throw residentNative.error;
        if (
          residentNative.status !== 0 ||
          !isDeepStrictEqual(
            JSON.parse(residentNative.stdout),
            JSON.parse(
              JSON.stringify({
                residentResult: expectedResident,
                document: check.expected.document,
              }),
            ),
          )
        )
          throw new TypeError(`Native resident fixture ${fixture.name} differs from TS.`);
        const residentVerified = await page.evaluate(
          (args) => {
            const verify = Reflect.get(globalThis, "verifyCanonicalSpikeResident");
            if (typeof verify !== "function")
              throw new TypeError("Resident verification is unavailable.");
            return verify(args);
          },
          {
            documentJson: JSON.stringify(check.document),
            opsJson: JSON.stringify(check.ops),
            expectedJson: JSON.stringify(expected),
          },
        );
        if (residentVerified !== true)
          throw new TypeError(`WASM resident fixture ${fixture.name} differs from TS.`);
        for (const arm of ["typescript", "wasm"] as const) {
          const verified = await page.evaluate(
            (args) => {
              const verify = Reflect.get(globalThis, "verifyCanonicalSpikeArm");
              if (typeof verify !== "function")
                throw new TypeError("Browser verification is unavailable.");
              return verify(args);
            },
            {
              arm,
              documentJson: JSON.stringify(check.document),
              opsJson: JSON.stringify(check.ops),
              expectedJson: JSON.stringify(expected),
            },
          );
          if (verified !== true)
            throw new TypeError(
              `Browser ${arm} fixture ${fixture.name} or inverse differs from TS.`,
            );
        }
      }
    } else {
      const nativeReplay = spawnSync(nativeBinary, [], {
        input: `${JSON.stringify({ benchmark: { documentJson: fixture.documentJson, opsJson: "[]", task: fixture.task, verify: true } })}\n`,
        encoding: "utf8",
      });
      if (nativeReplay.error) throw nativeReplay.error;
      const nativeModel: unknown = JSON.parse(nativeReplay.stdout);
      if (
        nativeReplay.status !== 0 ||
        typeof nativeModel !== "object" ||
        nativeModel === null ||
        !("verifiedModel" in nativeModel) ||
        !isDeepStrictEqual(nativeModel.verifiedModel, fixture.document)
      )
        throw new TypeError(`Native ${fixture.task} replay lost model fields.`);
      for (const arm of ["typescript", "wasm"] as const) {
        const verified = await page.evaluate(
          (args) => {
            const verify = Reflect.get(globalThis, "verifyCanonicalSpikeReplay");
            if (typeof verify !== "function")
              throw new TypeError("Replay verification is unavailable.");
            return verify(args);
          },
          { arm, documentJson: fixture.documentJson, task: fixture.task },
        );
        if (verified !== true)
          throw new TypeError(`Browser ${arm} ${fixture.task} replay lost model fields.`);
      }
    }
    record({
      type: "fixture-verified",
      name: fixture.name,
      pages: fixture.pages,
      inputBytes: new TextEncoder().encode(fixture.documentJson).byteLength,
      arms:
        fixture.task === "apply"
          ? [
              "typescript",
              "wasm",
              "native",
              "typescript-resident",
              "wasm-resident",
              "native-resident",
            ]
          : ["typescript", "wasm", "native"],
    });
    if (validateOnly) continue;
    if (fixture.task === "apply")
      await residentPage.evaluate((documentJson) => {
        const initialize = Reflect.get(globalThis, "initializeCanonicalSpikeResidents");
        if (typeof initialize !== "function")
          throw new TypeError("Resident initialization is unavailable.");
        initialize(documentJson);
      }, fixture.documentJson);
    for (let repetition = 0; repetition < 60; repetition += 1) {
      const arms =
        fixture.task === "apply"
          ? ([
              "typescript",
              "wasm",
              "native",
              "typescript-resident",
              "wasm-resident",
              "native-resident",
            ] as const)
          : (["typescript", "wasm", "native"] as const);
      // Rotation avoids assigning one arm systematically colder/earlier work.
      for (let index = 0; index < arms.length; index += 1) {
        const arm = arms.at((index + repetition) % arms.length);
        if (!arm) throw new TypeError("Missing benchmark arm.");
        const load1 = loadavg().at(0);
        if (load1 === undefined) throw new TypeError("Missing one-minute load.");
        if (quiet && load1 >= 3) {
          record({
            type: "load-refused",
            load1,
            pages: fixture.pages,
            name: fixture.name,
            repetition,
            arm,
          });
          throw new TypeError("Quiet-window load rose to 3 or above; stopping without retry.");
        }
        let sample: unknown;
        if (arm === "native" || arm === "native-resident") {
          const result = spawnSync(
            "/usr/bin/time",
            [platform() === "darwin" ? "-l" : "-v", nativeBinary],
            {
              input: `${JSON.stringify({ benchmark: { documentJson: fixture.documentJson, opsJson: fixture.opsJson, task: arm === "native-resident" ? "resident" : fixture.task, iterations: 1 } })}\n`,
              encoding: "utf8",
            },
          );
          if (result.error) throw result.error;
          if (result.status !== 0) throw new TypeError(`Native arm failed: ${result.stderr}`);
          const peak =
            platform() === "darwin"
              ? result.stderr.match(/(\d+)\s+maximum resident set size/u)?.at(1)
              : result.stderr.match(/Maximum resident set size \(kbytes\):\s*(\d+)/u)?.at(1);
          if (!peak) throw new TypeError("Native peak RSS was not recorded.");
          sample = {
            ...JSON.parse(result.stdout),
            peakRssBytes: Number(peak) * (platform() === "darwin" ? 1 : 1024),
          };
        } else if (arm === "typescript-resident" || arm === "wasm-resident") {
          sample = await residentPage.evaluate(
            (args) => {
              const run = Reflect.get(globalThis, "runCanonicalSpikeResidentSample");
              if (typeof run !== "function")
                throw new TypeError("Resident benchmark is unavailable.");
              return run(args);
            },
            { arm: arm === "wasm-resident" ? "wasm" : "typescript", opsJson: fixture.opsJson },
          );
        } else {
          sample = await page.evaluate(
            ({ arm: browserArm, documentJson, opsJson, task }) => {
              const run = Reflect.get(globalThis, "runCanonicalSpikeSample");
              if (typeof run !== "function")
                throw new TypeError("Browser benchmark is unavailable.");
              return run({ arm: browserArm, documentJson, opsJson, task });
            },
            {
              arm,
              documentJson: fixture.documentJson,
              opsJson: fixture.opsJson,
              task: fixture.task,
            },
          );
        }
        if (typeof sample !== "object" || sample === null || "status" in sample)
          throw new TypeError(`Benchmark arm did not apply: ${JSON.stringify(sample)}`);
        record({
          type: "sample",
          arm,
          pages: fixture.pages,
          name: fixture.name,
          task: fixture.task,
          repetition,
          warmup: repetition < 10,
          load1,
          sample,
          browserMemory:
            arm === "native" || arm === "native-resident"
              ? null
              : await (arm.endsWith("-resident") ? residentPage : page).evaluate(
                  () => Reflect.get(performance, "memory")?.usedJSHeapSize ?? null,
                ),
          time: `${new Date().toLocaleString("sv-SE", { timeZone: "Europe/Prague" })} CEST`,
        });
      }
    }
  }
} finally {
  await browser?.close();
  server.stop(true);
}
