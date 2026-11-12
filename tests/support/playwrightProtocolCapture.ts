import type { TestInfo } from "@playwright/test";
import {
  appendFileSync,
  createReadStream,
  createWriteStream,
  readdirSync,
  statSync,
} from "node:fs";
import { createRequire } from "node:module";
import { join } from "node:path";
import { pipeline } from "node:stream/promises";
import { format } from "node:util";

import { canonicalProtocolEvidence } from "../parity/canonicalProtocolEvidence";

/**
 * The one owner of Playwright protocol logging in this repository.
 *
 * A worker's stderr travels to the test runner, which keeps every chunk on the
 * test's result for the reporters until the run ends, and tracing copies it
 * into the trace as well. Protocol messages carry whole response bodies, DOM
 * snapshots and screencast frames, so they never take that path: each worker
 * writes payload-free evidence records to its own file in the capture
 * directory, and a failing test attaches its own slice of that file by path.
 * Every Playwright config passes through `withBoundedProtocolCapture`.
 */
export const PROTOCOL_CAPTURE_DIRECTORY = "FOLIO_PROTOCOL_CAPTURE_DIRECTORY";
export const PROTOCOL_NAMESPACE = "pw:protocol";
const EVIDENCE_SUFFIX = ".protocol.jsonl";

/** Environment that turns protocol capture on for one runner invocation. */
export const protocolCaptureEnvironment = (
  inherited: Record<string, string | undefined>,
  directory: string,
) => ({
  DEBUG: [inherited["DEBUG"], PROTOCOL_NAMESPACE].filter(Boolean).join(","),
  [PROTOCOL_CAPTURE_DIRECTORY]: directory,
});

/** This process's evidence file, when protocol capture is on. */
export const workerProtocolEvidencePath = () => {
  const directory = process.env[PROTOCOL_CAPTURE_DIRECTORY];
  return directory
    ? join(directory, `${process.pid}${EVIDENCE_SUFFIX}`)
    : undefined;
};

/** Every worker's evidence file in a capture directory. */
export const protocolEvidenceFiles = (directory: string) =>
  readdirSync(directory)
    .filter((name) => name.endsWith(EVIDENCE_SUFFIX))
    .sort()
    .map((name) => join(directory, name));

type DebugFactory = {
  (namespace: string): (message: string) => void;
  log: (this: unknown, ...args: unknown[]) => unknown;
};

const isDebugFactory = (value: unknown): value is DebugFactory =>
  typeof value === "function" &&
  "log" in value &&
  typeof value.log === "function";

/** The `debug` instance Playwright's protocol logger writes through. */
export const playwrightDebug = () => {
  const fromTest = createRequire(import.meta.url).resolve("@playwright/test");
  const fromRunner = createRequire(fromTest).resolve("playwright");
  const utils: unknown = createRequire(fromRunner)(
    "playwright-core/lib/utilsBundle",
  );
  const debug =
    utils !== null && typeof utils === "object" && "debug" in utils
      ? utils.debug
      : undefined;
  if (!isDebugFactory(debug))
    throw new TypeError("Playwright's debug logger has an unknown shape");
  return debug;
};

const appendEvidence = (path: string, line: string) => {
  // This boundary owns decoding failures; retain malformed evidence.
  try {
    const result = canonicalProtocolEvidence(line);
    switch (result.type) {
      case "output":
      case "discard":
        return;
      case "unparsed":
        appendFileSync(
          path,
          `${JSON.stringify({ unparsedLength: result.length })}\n`,
        );
        return;
      case "evidence":
        appendFileSync(
          path,
          `${JSON.stringify({ at: Date.now(), ...result.message })}\n`,
        );
        return;
    }
  } catch (error) {
    appendFileSync(
      path,
      `${JSON.stringify({ unparsedLength: line.length, error: error instanceof Error ? error.name : "UnknownError" })}\n`,
    );
  }
};

const installed = Symbol.for("folio.protocolCapture");

const installWorkerCapture = () => {
  // The runner only spawns workers; protocol traffic exists in workers alone.
  if (process.env["TEST_WORKER_INDEX"] === undefined) return;
  const path = workerProtocolEvidencePath();
  if (!path) return;
  const debug = playwrightDebug();
  if (installed in debug) return;
  const forward = debug.log;
  appendFileSync(path, "");
  debug.log = function (this: unknown, ...args: unknown[]) {
    // `debug` calls the logger with the debugger function itself as `this`.
    const namespace =
      typeof this === "function" && "namespace" in this
        ? this.namespace
        : undefined;
    if (namespace !== PROTOCOL_NAMESPACE) return forward.apply(this, args);
    appendEvidence(path, format(...args));
    return true;
  };
  Object.defineProperty(debug, installed, { value: true });
};

/** Every Playwright config passes through here so protocol logs stay out of the runner. */
export const withBoundedProtocolCapture = <Config>(config: Config) => {
  installWorkerCapture();
  return config;
};

/** Offset of this worker's evidence file, taken when a test starts. */
export const protocolEvidenceOffset = () => {
  const path = workerProtocolEvidencePath();
  return path ? statSync(path).size : 0;
};

/** Attach the evidence recorded since `offset` by path, so nothing is held in memory. */
export const attachProtocolEvidence = async (
  testInfo: TestInfo,
  name: string,
  offset: number,
) => {
  const path = workerProtocolEvidencePath();
  if (!path) return;
  const slice = testInfo.outputPath(`${name}.jsonl`);
  await pipeline(
    createReadStream(path, { start: offset }),
    createWriteStream(slice),
  );
  await testInfo.attach(name, {
    path: slice,
    contentType: "application/x-ndjson",
  });
};
