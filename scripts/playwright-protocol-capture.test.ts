/**
 * Playwright protocol logs never reach the test runner. The runner keeps every
 * stderr chunk a worker writes on that test's result until the run ends, so a
 * protocol log there grows with every traced action. One module owns protocol
 * capture and writes it to per-worker files; this file fails when a config,
 * workflow or script turns protocol logging on another way, or when the owner
 * stops keeping it out of the runner.
 */
import { expect, setDefaultTimeout, test } from "bun:test";
import { mkdtempSync, readFileSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join, relative } from "node:path";

import {
  expectExited,
  runChild,
  subprocessTestTimeout,
} from "../test/subprocess";
import {
  PROTOCOL_CAPTURE_DIRECTORY,
  protocolCaptureEnvironment,
  protocolEvidenceFiles,
} from "../tests/support/playwrightProtocolCapture";

setDefaultTimeout(subprocessTestTimeout());

const ROOT = join(import.meta.dir, "..");
const OWNER = "tests/support/playwrightProtocolCapture.ts";
const SELF = relative(ROOT, import.meta.path);
// The evidence parser and its test read the namespace; they never enable it.
const NAMESPACE_READERS = new Set([
  OWNER,
  SELF,
  "tests/parity/canonicalProtocolEvidence.ts",
  "tests/parity/canonicalProtocolEvidence.test.ts",
]);
const SCANNED = /\.(?:[cm]?[jt]sx?|json|ya?ml|sh|toml)$/u;
const PLAYWRIGHT_CONFIG = /(?:^|\/)playwright[^/]*\.config\.[cm]?[jt]s$/u;
const DEBUG_ENABLES_PLAYWRIGHT = /\bDEBUG\b[^\n]*\bpw:/u;
const WRAPPED_EXPORT = /export default withBoundedProtocolCapture\(/u;
const OWNER_IMPORT =
  /import \{ withBoundedProtocolCapture \} from "(?:\.\.?\/)+tests\/support\/playwrightProtocolCapture";/u;

type Source = { path: string; text: string };

/** Every way a source can put Playwright's protocol log back on the runner. */
const protocolCaptureViolations = (sources: readonly Source[]) => {
  const violations: string[] = [];
  for (const { path, text } of sources) {
    if (path === OWNER || path === SELF) continue;
    if (PLAYWRIGHT_CONFIG.test(path)) {
      if (!OWNER_IMPORT.test(text) || !WRAPPED_EXPORT.test(text))
        violations.push(
          `${path}: export the config through withBoundedProtocolCapture`,
        );
    }
    if (DEBUG_ENABLES_PLAYWRIGHT.test(text))
      violations.push(
        `${path}: sets DEBUG for Playwright; use protocolCaptureEnvironment`,
      );
    if (!NAMESPACE_READERS.has(path) && text.includes("pw:protocol"))
      violations.push(
        `${path}: names the protocol namespace; only ${OWNER} enables it`,
      );
  }
  return violations;
};

const trackedSources = (): Source[] => {
  const listed = expectExited(
    runChild(["git", "ls-files", "-z"], { cwd: ROOT }),
  );
  if (listed.exitCode !== 0)
    throw new TypeError(`git ls-files failed: ${listed.stderr}`);
  return listed.stdout
    .toString()
    .split("\0")
    .filter((path) => SCANNED.test(path))
    .map((path) => ({ path, text: readFileSync(join(ROOT, path), "utf8") }));
};

test("only the owner turns Playwright protocol logging on, and every config uses it", () => {
  const sources = trackedSources();
  expect(
    sources.filter(({ path }) => PLAYWRIGHT_CONFIG.test(path)).length,
  ).toBeGreaterThanOrEqual(3);
  expect(protocolCaptureViolations(sources)).toEqual([]);
});

test.each([
  {
    name: "a CI job enabling protocol logs",
    source: {
      path: ".github/workflows/ci.yml",
      text: "      - run: bun run test:interactions\n        env:\n          DEBUG: pw:protocol\n",
    },
    violation: "sets DEBUG for Playwright",
  },
  {
    name: "a CI job enabling every Playwright namespace",
    source: {
      path: ".github/workflows/ci.yml",
      text: "        run: DEBUG=pw:* bunx playwright test\n",
    },
    violation: "sets DEBUG for Playwright",
  },
  {
    name: "a runner script appending the namespace",
    source: {
      path: "scripts/canonical-interactions.ts",
      text: 'const env = { DEBUG: [process.env["DEBUG"], "pw:protocol"].join(",") };\n',
    },
    violation: "sets DEBUG for Playwright",
  },
  {
    name: "a package script enabling protocol logs",
    source: {
      path: "package.json",
      text: '{ "scripts": { "test:interactions": "DEBUG=pw:protocol bunx playwright test" } }\n',
    },
    violation: "sets DEBUG for Playwright",
  },
  {
    name: "a config that bypasses the owner",
    source: {
      path: "packages/editor-web/playwright.config.ts",
      text: 'import { defineConfig } from "@playwright/test";\n\nexport default defineConfig({});\n',
    },
    violation: "export the config through withBoundedProtocolCapture",
  },
  {
    name: "a helper that names the namespace",
    source: {
      path: "tests/visual/protocolDebug.ts",
      text: 'process.env["NODE_DEBUG_NAMESPACES"] = "pw:protocol";\n',
    },
    violation: "names the protocol namespace",
  },
])("the guard refuses $name", ({ source, violation }) => {
  const found = protocolCaptureViolations([source]);
  expect(found.some((message) => message.includes(violation))).toBe(true);
});

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

const records = (value: unknown, key: string): unknown[] => {
  if (!isRecord(value)) return [];
  const children = value[key];
  return Array.isArray(children) ? children : [];
};

/** Bytes the runner kept on test results, read from its JSON report. */
const retainedStderrBytes = (report: unknown) => {
  let bytes = 0;
  let results = 0;
  const visit = (suite: unknown) => {
    for (const spec of records(suite, "specs"))
      for (const entry of records(spec, "tests"))
        for (const result of records(entry, "results")) {
          results += 1;
          for (const chunk of records(result, "stderr")) {
            if (!isRecord(chunk)) continue;
            const text = chunk["text"] ?? chunk["buffer"];
            if (typeof text === "string") bytes += Buffer.byteLength(text);
          }
        }
    for (const child of records(suite, "suites")) visit(child);
  };
  for (const suite of records(report, "suites")) visit(suite);
  return { bytes, results };
};

const FLOOD_MESSAGES = 64;
const FLOOD_PAYLOAD_BYTES = 64 * 1024;
const FLOOD_BYTES = FLOOD_MESSAGES * FLOOD_PAYLOAD_BYTES;
// Far below one message, so even a single leaked protocol line fails.
const RETAINED_CAP_BYTES = FLOOD_PAYLOAD_BYTES / 2;
const EVIDENCE_RECORD_CAP_BYTES = 1024;

const runFlood = (environment: Record<string, string>) => {
  const output = mkdtempSync(join(tmpdir(), "folio-protocol-flood-"));
  try {
    const inherited = Object.fromEntries(
      Object.entries(process.env).filter(
        ([name]) => name !== PROTOCOL_CAPTURE_DIRECTORY && name !== "DEBUG",
      ),
    );
    const run = expectExited(
      runChild(
        [
          "node",
          join(ROOT, "node_modules/.bin/playwright"),
          "test",
          "-c",
          join(
            import.meta.dir,
            "__fixtures__/protocol-capture/playwright.config.ts",
          ),
          `--output=${output}`,
        ],
        {
          cwd: ROOT,
          env: {
            ...inherited,
            FOLIO_PROTOCOL_FLOOD_MESSAGES: String(FLOOD_MESSAGES),
            FOLIO_PROTOCOL_FLOOD_PAYLOAD_BYTES: String(FLOOD_PAYLOAD_BYTES),
            ...environment,
          },
        },
      ),
    );
    const report: unknown = JSON.parse(run.stdout);
    expect(run.exitCode, run.stderr).toBe(0);
    return retainedStderrBytes(report);
  } finally {
    rmSync(output, { recursive: true, force: true });
  }
};

test("a protocol flood without the owner lands on the runner's test result", () => {
  const retained = runFlood({ DEBUG: "pw:protocol" });
  expect(retained.results).toBe(1);
  expect(retained.bytes).toBeGreaterThan(FLOOD_BYTES);
});

test("with the owner the runner keeps none of the flood and the worker keeps payload-free evidence", () => {
  const directory = mkdtempSync(join(tmpdir(), "folio-protocol-capture-"));
  try {
    const retained = runFlood(protocolCaptureEnvironment({}, directory));
    expect(retained.results).toBe(1);
    expect(retained.bytes).toBeLessThan(RETAINED_CAP_BYTES);
    const lines = protocolEvidenceFiles(directory).flatMap((file) =>
      readFileSync(file, "utf8").split("\n").filter(Boolean),
    );
    expect(lines).toHaveLength(FLOOD_MESSAGES);
    for (const line of lines) {
      expect(Buffer.byteLength(line)).toBeLessThan(EVIDENCE_RECORD_CAP_BYTES);
      expect(JSON.parse(line)).toMatchObject({
        sessionId: "flood",
        method: "Runtime.evaluate",
      });
    }
  } finally {
    rmSync(directory, { recursive: true, force: true });
  }
});
