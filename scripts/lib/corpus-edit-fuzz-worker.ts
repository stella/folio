import { spawnSync } from "node:child_process";
import { createHash } from "node:crypto";
import { mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";

import { unzipDocx } from "@stll/folio-core/docx/unzip";

import { openReviewer, toArrayBuffer } from "../../test/consumer-scenarios/support/documents.ts";
import {
  assertRequestedOutcome,
  capture,
  resolvedState,
  rowsOf,
} from "../../test/consumer-scenarios/support/oracle.ts";
import {
  coreBatch,
  type Mode,
  type Operation,
  randomOperation,
} from "../../test/consumer-scenarios/support/operations.ts";
import { createRandom } from "../../test/consumer-scenarios/support/random.ts";

const TYPES = [
  "replaceInBlock",
  "insertAfterBlock",
  "deleteBlock",
  "splitBlock",
  "mergeBlockWithNext",
  "formatRange",
  "commentOnRange",
  "insertTableRow",
  "deleteTableRow",
] as const;

export type EditStep = { mode: Mode; operation: Operation };
export type EditFailure = {
  class: string;
  signature: string;
  expected: string;
  observed: string;
  operations: EditStep[];
};
export type EditWorkerResult = {
  status: "parsed" | "unparsed";
  attempts: number;
  failures: EditFailure[];
};

const messageOf = (error: unknown): string =>
  error instanceof Error ? error.message : String(error);
const sdkDll = path.resolve(
  import.meta.dir,
  "../../packages/core/scripts/differential/dotnet/bin/Release/net8.0/OpenXmlProjector.dll",
);

const sdkErrors = (bytes: Uint8Array): string[] => {
  const directory = mkdtempSync(path.join(tmpdir(), "folio-edit-fuzz-"));
  try {
    const file = path.join(directory, "document.docx");
    writeFileSync(file, bytes);
    const result = spawnSync("dotnet", [sdkDll, "validate-all", file], {
      encoding: "utf8",
      timeout: 30_000,
      maxBuffer: 64 * 1024 * 1024,
    });
    if (result.error || result.status !== 0) {
      throw new Error(
        `Open XML SDK validator failed: ${result.error?.message ?? result.stderr.trim()}`,
      );
    }
    const report: unknown = JSON.parse(result.stdout);
    if (
      typeof report !== "object" ||
      report === null ||
      !("errors" in report) ||
      !Array.isArray(report.errors)
    ) {
      throw new Error("Open XML SDK validator returned an invalid report");
    }
    return report.errors.filter((error): error is string => typeof error === "string");
  } finally {
    rmSync(directory, { recursive: true, force: true });
  }
};

const partsOf = async (bytes: Uint8Array): Promise<Map<string, Uint8Array>> => {
  const raw = await unzipDocx(toArrayBuffer(bytes), { extractAllXml: true });
  return new Map(
    await Promise.all(
      Object.entries(raw.originalZip.files)
        .filter(([, file]) => !file.dir)
        .map(async ([name, file]) => [name, await file.async("uint8array")] as const),
    ),
  );
};

const sameBytes = (left: Uint8Array, right: Uint8Array): boolean =>
  left.length === right.length && left.every((byte, index) => byte === right[index]);

const allowedParts = (operation: Operation): ReadonlySet<string> => {
  const common = ["word/document.xml", "docProps/core.xml"];
  switch (operation.type) {
    case "commentOnRange":
      return new Set([
        ...common,
        "word/comments.xml",
        "word/_rels/document.xml.rels",
        "[Content_Types].xml",
      ]);
    default:
      return new Set([
        ...common,
        ...(operation["numbering"]
          ? ["word/numbering.xml", "word/_rels/document.xml.rels", "[Content_Types].xml"]
          : []),
      ]);
  }
};

const localityProblems = async (
  control: Uint8Array,
  edited: Uint8Array,
  operation: Operation,
): Promise<string[]> => {
  const before = await partsOf(control);
  const after = await partsOf(edited);
  const allowed = allowedParts(operation);
  const changed: string[] = [];
  for (const name of new Set([...before.keys(), ...after.keys()])) {
    if (allowed.has(name)) continue;
    const left = before.get(name);
    const right = after.get(name);
    if (!left || !right || !sameBytes(left, right)) changed.push(name);
  }
  return changed.sort();
};

const compareSdk = (original: readonly string[], saved: readonly string[]): string[] => {
  const known = new Map<string, number>();
  for (const error of original) known.set(error, (known.get(error) ?? 0) + 1);
  return saved.filter((error) => {
    const remaining = known.get(error) ?? 0;
    if (remaining === 0) return true;
    known.set(error, remaining - 1);
    return false;
  });
};

const saveAndReopen = async (
  reviewer: Awaited<ReturnType<typeof openReviewer>>,
): Promise<Uint8Array> => {
  const before = reviewer.getContent().map(({ text, kind }) => ({ text, kind }));
  const saved = new Uint8Array(await reviewer.toBuffer());
  const reopened = await openReviewer(saved);
  const after = reopened.getContent().map(({ text, kind }) => ({ text, kind }));
  if (JSON.stringify(before) !== JSON.stringify(after)) {
    throw new Error("saved package reopens to different blocks");
  }
  return saved;
};

const runSteps = async (
  original: Uint8Array,
  steps: readonly EditStep[],
  originalSdk: readonly string[] | null,
): Promise<EditFailure | null> => {
  const reviewer = await openReviewer(original);
  const control = new Uint8Array(await reviewer.toBuffer());
  for (const [index, { mode, operation }] of steps.entries()) {
    try {
      const pre = await capture(reviewer, mode);
      const batch = coreBatch([operation], mode);
      let result: ReturnType<typeof reviewer.applyDocumentOperations>;
      try {
        result = reviewer.applyDocumentOperations(batch as never);
      } catch (error) {
        return failure(
          "throw",
          operation.type,
          "a typed issue or applied receipt",
          messageOf(error),
          steps.slice(0, index + 1),
        );
      }
      const applied = result.applied.length > 0;
      try {
        const gaps = await assertRequestedOutcome(
          reviewer,
          pre,
          { applied: applied ? [operation] : [] },
          `edit ${index}`,
        );
        if (applied && gaps.length > 0) {
          return failure(
            "oracle-gap",
            operation.type,
            "requested outcome is modelled",
            gaps.join("; "),
            steps.slice(0, index + 1),
          );
        }
      } catch (error) {
        return failure(
          "outcome",
          operation.type,
          "requested outcome after save and reopen",
          messageOf(error),
          steps.slice(0, index + 1),
        );
      }
      let saved: Uint8Array;
      try {
        saved = await saveAndReopen(reviewer);
      } catch (error) {
        return failure(
          "reopen",
          operation.type,
          "saved package reopens to the edited view",
          messageOf(error),
          steps.slice(0, index + 1),
        );
      }
      if (!applied) continue;
      if (index === 0 && mode === "direct") {
        const changed = await localityProblems(control, saved, operation);
        if (changed.length > 0) {
          return failure(
            "locality",
            operation.type,
            "unrelated package parts byte-identical to control save",
            changed.join(", "),
            steps.slice(0, index + 1),
          );
        }
      }
      if (mode === "tracked-changes") {
        const rejected = await resolvedState(saved, "reject");
        const before = await resolvedState(original, "reject");
        if (
          JSON.stringify(rowsOf(await openReviewer(rejected.bytes)).map(({ text }) => text)) !==
          JSON.stringify(before.rows.map(({ text }) => text))
        ) {
          return failure(
            "reject",
            operation.type,
            "reject-all equals original text",
            "rejected text differs",
            steps.slice(0, index + 1),
          );
        }
      }
      if (originalSdk) {
        const newErrors = compareSdk(originalSdk, sdkErrors(saved));
        if (newErrors.length > 0) {
          return failure(
            "schema",
            operation.type,
            "no new Open XML SDK errors",
            newErrors.slice(0, 3).join(" | "),
            steps.slice(0, index + 1),
          );
        }
      }
    } catch (error) {
      return failure(
        "throw",
        operation.type,
        "operation and checks complete without throwing",
        messageOf(error),
        steps.slice(0, index + 1),
      );
    }
  }
  return null;
};

const failure = (
  kind: string,
  type: string,
  expected: string,
  observed: string,
  operations: readonly EditStep[],
): EditFailure => {
  let category = "result";
  if (kind === "outcome") {
    if (observed.includes("block texts differ")) category = "text";
    else if (observed.includes("comment")) category = "comment";
    else category = "other";
  } else if (kind === "locality") {
    category = "parts";
  }
  const digest = createHash("sha256").update(observed).digest("hex");
  return {
    class: kind,
    signature: `${kind}:${type}:${category}`,
    expected,
    observed: `${category} mismatch (diagnostic SHA-256 ${digest})`,
    operations: [...operations],
  };
};

const shrink = async (
  original: Uint8Array,
  found: EditFailure,
  originalSdk: readonly string[] | null,
): Promise<EditFailure> => {
  const count = found.operations.length;
  const masks = Array.from({ length: (1 << count) - 2 }, (_, index) => index + 1).sort(
    (left, right) =>
      left.toString(2).replaceAll("0", "").length - right.toString(2).replaceAll("0", "").length ||
      left - right,
  );
  for (const mask of masks) {
    const candidate = found.operations.filter((_, index) => (mask & (1 << index)) !== 0);
    try {
      // oxlint-disable-next-line no-await-in-loop -- try shorter subsequences first, preserving request order
      const replay = await runSteps(original, candidate, originalSdk);
      if (replay?.signature === found.signature) return replay;
    } catch {
      // A removed setup operation can make a later target inapplicable.
    }
  }
  return found;
};

export const runEditWorker = async (
  pathToDocx: string,
  seed: number,
  skipSdk: boolean,
): Promise<EditWorkerResult> => {
  const original = new Uint8Array(await Bun.file(pathToDocx).arrayBuffer());
  try {
    await openReviewer(original);
  } catch {
    return { status: "unparsed", attempts: 0, failures: [] };
  }
  const failures: EditFailure[] = [];
  let originalSdk: string[] | null = null;
  if (!skipSdk) {
    try {
      originalSdk = sdkErrors(original);
    } catch (error) {
      failures.push(
        failure(
          "sdk-original",
          "document",
          "Open XML SDK validates the original package",
          messageOf(error),
          [],
        ),
      );
    }
  }
  let attempts = 0;
  for (const mode of ["direct", "tracked-changes"] as const) {
    const random = createRandom(seed ^ (mode === "direct" ? 0x1234abcd : 0x5678ef01));
    const reviewer = await openReviewer(original);
    const steps: EditStep[] = [];
    for (let index = 0; index < 3; index += 1) {
      const operation = randomOperation(reviewer.getContent(), mode, random, TYPES);
      if (!operation) break;
      steps.push({ mode, operation });
      attempts += 1;
      try {
        reviewer.applyDocumentOperations(coreBatch([operation], mode) as never);
      } catch {
        break;
      }
    }
    if (steps.length > 0) {
      const found = await runSteps(original, steps, originalSdk);
      if (found) failures.push(await shrink(original, found, originalSdk));
    }
  }
  return { status: "parsed", attempts, failures };
};

export const replayEditCase = async (
  pathToDocx: string,
  steps: readonly EditStep[],
  skipSdk: boolean,
): Promise<EditFailure | null> => {
  const original = new Uint8Array(await Bun.file(pathToDocx).arrayBuffer());
  return runSteps(original, steps, skipSdk ? null : sdkErrors(original));
};
