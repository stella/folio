import { expect, spyOn, test } from "bun:test";
import { Result } from "better-result";
import type { Command } from "prosemirror-state";
import { getCanonicalCommandIntents } from "../prosemirror/canonicalCommands";
import {
  canonicalRefusalCaseId,
  declareCanonicalRefusalCases,
} from "../../../../test/canonical-conformance-refusals";
import * as canonicalStructure from "./canonicalStructure";
import { CanonicalSessionError } from "./canonicalSession";
import { CANONICAL_GAP } from "../types/canonicalCapabilities";
import { DOCUMENT_SHAPES, documentShape } from "../__tests__/documentShapes";
import { EDITOR_MODES } from "../__tests__/editorHarness";
import { CONFORMANCE_OPERATIONS, runConformanceCase } from "../__tests__/editorCommandConformance";

const CASE = {
  shape: "plain-markdown",
  operation: "command:toggleBold",
  placement: "word",
} as const;
const DECLARED_CASES = EDITOR_MODES.map((mode) => ({ ...CASE, mode }));

test("canonical refusal table covers exactly the declared concrete cases", () => {
  const rows = declareCanonicalRefusalCases(DECLARED_CASES);
  expect([...rows.keys()].toSorted()).toEqual(
    DECLARED_CASES.map(canonicalRefusalCaseId).toSorted(),
  );
  expect([...rows.values()]).toEqual([[], []]);
  expect(() => declareCanonicalRefusalCases([...DECLARED_CASES, ...DECLARED_CASES])).toThrow(
    "Duplicate canonical refusal case",
  );
});

test.each(EDITOR_MODES)(
  "table activation contracts match the declared fixture features in %s",
  (mode) => {
    const cases = DOCUMENT_SHAPES.map(({ id }) => ({ ...CASE, shape: id, mode }));
    const rows = declareCanonicalRefusalCases(cases);
    const activationShapes = cases
      .filter((key) =>
        rows
          .get(canonicalRefusalCaseId(key))
          ?.some(({ gap }) => gap === CANONICAL_GAP.tableActivation),
      )
      .map(({ shape }) => shape);
    const tableShapes = DOCUMENT_SHAPES.filter(
      ({ features }) =>
        features.includes("table") ||
        features.includes("nested-table") ||
        features.includes("merged-cells"),
    ).map(({ id }) => id);
    expect(activationShapes.toSorted()).toEqual(tableShapes.toSorted());
    expect(rows.get(canonicalRefusalCaseId({ ...CASE, shape: "tables", mode }))).toEqual([
      {
        id: "table-session-activation",
        gap: CANONICAL_GAP.tableActivation,
        message: "Canonical sessions cannot activate documents containing tables.",
      },
    ]);
    expect(rows.get(canonicalRefusalCaseId({ ...CASE, mode }))).toEqual([]);
  },
);

// Inject the same ledger ids the old planner preview automatically allowed.
// The control proves bold is supported; the ablation must fail the conformance oracle.
test.each([
  {
    gap: CANONICAL_GAP.trackedHyperlinkResolution,
    message: "Hyperlink suggestions require serializable wrapper review provenance.",
  },
  {
    gap: CANONICAL_GAP.storyContentProjection,
    message: "The operations produce unsupported canonical story content.",
  },
])("supported plain-content commands cannot acquire a $gap allowance", async ({ gap, message }) => {
  const operation = CONFORMANCE_OPERATIONS.find(({ id }) => id === CASE.operation);
  if (!operation) throw new TypeError("Refusal oracle control command is absent");
  const options = {
    shape: documentShape(CASE.shape),
    operation,
    placement: CASE.placement,
    refusalCases: declareCanonicalRefusalCases(DECLARED_CASES),
  };
  const control = await runConformanceCase(options);
  expect(control?.runs).toEqual({ editing: "changed", suggesting: "changed" });
  expect(control?.violations).toEqual([]);
  const planner = spyOn(canonicalStructure, "prepareCanonicalCommands").mockImplementation(() =>
    Result.err(new CanonicalSessionError({ gap, message, reason: "refused" })),
  );
  try {
    const ablated = await runConformanceCase(options);
    expect(ablated?.runs).toEqual({ editing: "refused", suggesting: "refused" });
    expect(ablated?.violations).toEqual(
      EDITOR_MODES.map((mode) => ({ kind: "silent-refusal", mode, detail: `${gap}: ${message}` })),
    );
    expect(ablated?.refusals.map(({ expectation }) => expectation)).toEqual([
      "unexpected",
      "unexpected",
    ]);
  } finally {
    planner.mockRestore();
  }
});

test("a supported command losing its descriptor fails standing conformance in both modes", async () => {
  const operation = CONFORMANCE_OPERATIONS.find(({ id }) => id === CASE.operation);
  if (!operation) throw new TypeError("Descriptor-loss control command is absent");
  const refusalCases = declareCanonicalRefusalCases(DECLARED_CASES);
  expect([...refusalCases.values()]).toEqual([[], []]);
  const options = {
    shape: documentShape(CASE.shape),
    operation,
    placement: CASE.placement,
    refusalCases,
  };
  const control = await runConformanceCase(options);
  expect(control?.runs).toEqual({ editing: "changed", suggesting: "changed" });
  expect(control?.violations).toEqual([]);
  expect(control?.refusals).toEqual([]);
  let descriptorLossControls = 0;
  const ablated = await runConformanceCase({
    ...options,
    operation: {
      ...operation,
      run: ({ view }) => {
        if (view.authority !== "canonical")
          throw new TypeError("Descriptor-loss case lost canonical authority");
        const command = view.commandManager.requireCommand("toggleBold")();
        const withoutDescriptor: Command = (state, dispatch, editorView) =>
          command(state, dispatch, editorView);
        // Preserve the exact command body; only its WeakMap metadata is lost.
        expect(getCanonicalCommandIntents(command, view.state)?.length).toBeGreaterThan(0);
        expect(getCanonicalCommandIntents(withoutDescriptor, view.state)).toBeUndefined();
        descriptorLossControls++;
        return view.execute(withoutDescriptor);
      },
    },
  });
  expect(descriptorLossControls).toBe(EDITOR_MODES.length);
  expect(ablated?.runs).toEqual({ editing: "refused", suggesting: "refused" });
  expect(ablated?.violations).toEqual(
    EDITOR_MODES.map((mode) => ({
      kind: "silent-refusal",
      mode,
      detail: `${CANONICAL_GAP.dispatch}: Unclassified native text is unavailable in this session.`,
    })),
  );
  expect(ablated?.refusals.map(({ expectation }) => expectation)).toEqual([
    "unexpected",
    "unexpected",
  ]);
});

test("unsupported note insertion has an explicit ledger-keyed conformance contract", async () => {
  const key = { ...CASE, operation: "command:insertFootnote", placement: "caret-middle" } as const;
  const operation = CONFORMANCE_OPERATIONS.find(({ id }) => id === key.operation);
  if (!operation) throw new TypeError("Explicit descriptor-gap case is absent");
  const refusalCases = declareCanonicalRefusalCases(EDITOR_MODES.map((mode) => ({ ...key, mode })));
  const row = {
    id: CANONICAL_GAP.commands,
    gap: CANONICAL_GAP.dispatch,
    message: "Unclassified native text is unavailable in this session.",
  };
  expect([...refusalCases.values()]).toEqual([[row], [row]]);
  const result = await runConformanceCase({
    shape: documentShape(key.shape),
    operation,
    placement: key.placement,
    refusalCases,
  });
  expect(result?.runs).toEqual({ editing: "refused", suggesting: "refused" });
  expect(result?.violations).toEqual([]);
  expect(result?.refusals).toEqual(
    EDITOR_MODES.map(() => ({
      gap: row.gap,
      message: row.message,
      row: row.id,
      expectation: "declared",
    })),
  );
});
