/**
 * Editor-command conformance: every command, key binding, typed marker, paste
 * and host flow × every document shape × editing and suggesting mode.
 *
 * Tiers:
 * - default (CI): each operation at the placements it names;
 * - full: every operation at every selection placement —
 *   `FOLIO_CONFORMANCE=full bun test src/__tests__/editorCommandConformance.test.ts`
 *   (or `bun run test:conformance:full` in this package).
 *
 * `FOLIO_CONFORMANCE_FILTER=<regex>` narrows the run to matching case ids
 * (`shape › operation @ placement`).
 */

import { afterAll, describe, expect, test } from "bun:test";
import { appendFileSync } from "node:fs";

import { DOCUMENT_SHAPES } from "./documentShapes";
import {
  CONFORMANCE_OPERATIONS,
  KEY_BINDING_OPERATIONS,
  REGISTRY_COMMAND_OPERATIONS,
  runConformanceCase,
} from "./editorCommandConformance";
import type { Violation } from "./editorCommandConformance";
import { gapApplies, gapCovers, KNOWN_CONFORMANCE_GAPS } from "./editorCommandConformance.known";
import type { ConformanceCaseKey, KnownConformanceGap } from "./editorCommandConformance.known";
import { EDITOR_MODES, harnessManager, SELECTION_PLACEMENTS } from "./editorHarness";
import { createStarterKit } from "../prosemirror/extensions/StarterKit";

const FULL_TIER = process.env["FOLIO_CONFORMANCE"] === "full";
const FILTER = process.env["FOLIO_CONFORMANCE_FILTER"]
  ? new RegExp(process.env["FOLIO_CONFORMANCE_FILTER"], "u")
  : null;

/** Append every case's result as a JSON line here, for triage. */
const REPORT_PATH = process.env["FOLIO_CONFORMANCE_REPORT"];

const caseId = ({ shape, operation, placement }: ConformanceCaseKey): string =>
  `${shape} › ${operation} @ ${placement}`;

const CASES: ConformanceCaseKey[] = DOCUMENT_SHAPES.flatMap((shape) =>
  CONFORMANCE_OPERATIONS.flatMap((operation) =>
    (FULL_TIER ? SELECTION_PLACEMENTS : operation.placements).map((placement) => ({
      shape: shape.id,
      operation: operation.id,
      placement,
    })),
  ),
).filter((key) => FILTER === null || FILTER.test(caseId(key)));

const gapUsage = new Map<KnownConformanceGap, { applied: number; covered: number }>(
  KNOWN_CONFORMANCE_GAPS.map((gap) => [gap, { applied: 0, covered: 0 }]),
);

let executedCases = 0;

const describeViolation = (violation: Violation): string =>
  `[${violation.mode}] ${violation.kind}: ${violation.detail}`;

describe("editor command conformance: coverage", () => {
  test("every registry command is driven or excluded with a reason", () => {
    const registered = Object.keys(harnessManager().getCommands()).toSorted();
    const catalogued = Object.keys(REGISTRY_COMMAND_OPERATIONS).toSorted();
    expect(registered.filter((name) => !catalogued.includes(name))).toEqual([]);
    expect(catalogued.filter((name) => !registered.includes(name))).toEqual([]);
  });

  test("every key binding is driven or excluded with a reason", () => {
    const schema = harnessManager().getSchema();
    const bindings = new Set<string>();
    for (const extension of createStarterKit()) {
      for (const binding of Object.keys(
        extension.onSchemaReady({ schema }).keyboardShortcuts ?? {},
      )) {
        bindings.add(binding);
      }
    }
    const catalogued = Object.keys(KEY_BINDING_OPERATIONS);
    expect([...bindings].filter((binding) => !catalogued.includes(binding)).toSorted()).toEqual([]);
    // The base keymap binds the macOS motion aliases only on macOS.
    const onThisPlatform = catalogued.filter((binding) => {
      const entry = KEY_BINDING_OPERATIONS[binding];
      return !(entry && "macOSOnly" in entry && entry.macOSOnly === true);
    });
    expect(onThisPlatform.filter((binding) => !bindings.has(binding)).toSorted()).toEqual([]);
  });

  test("operation ids are unique", () => {
    const ids = CONFORMANCE_OPERATIONS.map((operation) => operation.id);
    expect(ids.filter((id, index) => ids.indexOf(id) !== index)).toEqual([]);
  });
});

describe("editor command conformance", () => {
  const shapesById = new Map(DOCUMENT_SHAPES.map((shape) => [shape.id, shape]));
  const operationsById = new Map(
    CONFORMANCE_OPERATIONS.map((operation) => [operation.id, operation]),
  );

  test.each(CASES.map((key) => [caseId(key), key] as const))(
    "%s",
    async (_id, key) => {
      const shape = shapesById.get(key.shape);
      const operation = operationsById.get(key.operation);
      if (!shape || !operation) {
        throw new Error(`Unknown case ${caseId(key)}`);
      }
      const result = await runConformanceCase(shape, operation, key.placement);
      if (!result) {
        return;
      }
      executedCases += EDITOR_MODES.length;

      for (const gap of KNOWN_CONFORMANCE_GAPS) {
        if (gapApplies(gap, key)) {
          const usage = gapUsage.get(gap);
          if (usage) {
            usage.applied += 1;
          }
        }
      }
      const unexpected = result.violations.filter((violation) => {
        const gap = KNOWN_CONFORMANCE_GAPS.find((candidate) =>
          gapCovers(candidate, key, violation.kind, violation.mode),
        );
        if (!gap) {
          return true;
        }
        const usage = gapUsage.get(gap);
        if (usage) {
          usage.covered += 1;
        }
        return false;
      });
      if (REPORT_PATH) {
        appendFileSync(REPORT_PATH, `${JSON.stringify({ ...key, ...result, unexpected })}\n`);
      }
      expect(unexpected.map(describeViolation)).toEqual([]);
    },
    60_000,
  );

  test("no known gap is stale", () => {
    const stale = [...gapUsage.entries()]
      .filter(
        ([gap, usage]) =>
          usage.applied > 0 && usage.covered === 0 && (gap.tier !== "full" || FULL_TIER),
      )
      .map(([gap]) => `${gap.issue === undefined ? "" : `#${gap.issue} `}${gap.reason}`);
    expect(stale).toEqual([]);
  });

  afterAll(() => {
    // One line for the CI log: how much of the matrix ran.
    console.info(
      `editor command conformance: ${executedCases} cases (${CONFORMANCE_OPERATIONS.length} operations × ${DOCUMENT_SHAPES.length} shapes × ${EDITOR_MODES.length} modes, ${FULL_TIER ? "full" : "default"} tier)`,
    );
  });
});
