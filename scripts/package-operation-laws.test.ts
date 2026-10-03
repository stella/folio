/** Package byte and scope laws extend the corpus oracle to every operation kind in unit CI. */
import { expect, test } from "bun:test";
import fc from "fast-check";
import { TaggedError } from "better-result";
import { assertKnownProperty, assertProperty, propertyTestTimeout } from "../test/property-testing";
import {
  OPERATION_LAW_DISPOSITIONS,
  type OperationLawDisposition,
} from "../test/operation-law-dispositions";
import {
  inverseSequenceFailures,
  serializedInverseStepFailures,
} from "./lib/corpus-invariants/op-inverse";
import {
  localityStepFailures,
  serializedLocalityStepFailures,
} from "./lib/corpus-invariants/op-locality";
import {
  exactOpModel,
  sameOpModel,
  serializeOpDocument,
  serializedOpParts,
} from "./lib/corpus-invariants/op-sequences";
import { createDocx } from "@stll/folio-core/docx/rezip";
import { parseDocx } from "@stll/folio-core/docx/parser";
import { normalizeForOps } from "../packages/docx-core/src/ops/contract";
import { applyDocumentOps } from "../packages/docx-core/src/ops/apply";
import { DOCUMENT_OP_TYPES } from "../packages/docx-core/src/ops/types";
import { failureMarker } from "../test/consumer-scenarios/support/failure-fingerprints";
import { opSeedArbitrary } from "../packages/docx-core/src/ops/__tests__/documentArbitraries";
import {
  generatedCaseFor,
  GENERATED_PACKAGE_OP_KINDS,
  GENERATED_PACKAGE_STORIES,
  packageDocumentArbitrary,
} from "../packages/docx-core/src/ops/__tests__/packageOperationArbitraries";

class OperationPackageLawError extends TaggedError("OperationPackageLawError")<{
  message: string;
}> {}

const LAW_TITLE = "every operation preserves package inverse and declared scope laws";
const LAW_KEY = `scripts/package-operation-laws.test.ts::${LAW_TITLE}`;

const knownFailures = (disposition: OperationLawDisposition, kind: string) => {
  if (disposition === "holds") return [];
  const issues = "knownIssue" in disposition ? [disposition] : disposition;
  return issues.map(({ knownIssue, fingerprint }) => ({
    family: knownIssue,
    fingerprint,
    matches: (value: readonly unknown[]) => value.at(0) === kind,
  }));
};

const unexpectedFailure = (
  kind: keyof typeof OPERATION_LAW_DISPOSITIONS,
  failures: readonly string[],
) => {
  const expected = knownFailures(OPERATION_LAW_DISPOSITIONS[kind], kind);
  return failures.find((message) => {
    const marker = failureMarker({
      test: LAW_KEY,
      seed: 0,
      repro: "",
      failure: new OperationPackageLawError({ message }),
    });
    return !expected.some(({ fingerprint }) => fingerprint === marker.fingerprint);
  });
};

test("a recorded inverse symptom cannot hide a new scope violation", () => {
  const inverse =
    "joinBlocks inverse changed the original serialized package parts: word/header1.xml";
  const scope = "joinBlocks changed unrelated serialized part: word/comments.xml";
  expect(unexpectedFailure("joinBlocks", [inverse, scope])).toBe(scope);
  expect(unexpectedFailure("insertText", [inverse])).toBe(inverse);
});

test(
  "every operation preserves package inverse and declared scope laws",
  async () => {
    expect(new Set(GENERATED_PACKAGE_OP_KINDS)).toEqual(new Set(Object.values(DOCUMENT_OP_TYPES)));
    const property = fc.asyncProperty(
      fc.constantFrom(...GENERATED_PACKAGE_OP_KINDS),
      packageDocumentArbitrary,
      opSeedArbitrary,
      fc.constantFrom(...GENERATED_PACKAGE_STORIES),
      async (kind, document, seed, story) => {
        const bytes = await createDocx(structuredClone(document));
        const packaged = normalizeForOps(await parseDocx(bytes, { preloadFonts: false }));
        const generated = generatedCaseFor({ document: packaged, seed, story, kind });
        expect(generated.op.type).toBe(kind);
        const before = generated.document;
        const originalModel = exactOpModel(before);
        const originalXml = serializeOpDocument(before);
        const control = await serializedOpParts(before);
        const applied = applyDocumentOps(before, [generated.op]);
        if (applied.isErr())
          throw new OperationPackageLawError({
            message: `${kind} refused generated case: ${applied.error.reason}: ${applied.error.message}`,
          });
        const edit = applied.value;
        const step = { before, op: generated.op, edit };
        const failures = inverseSequenceFailures({
          original: before,
          originalModel,
          originalXml,
          document: edit.document,
          steps: [step],
          inverse: [...edit.inverse],
          mutations: sameOpModel(before, originalModel) ? [] : [kind],
          refusals: [],
        });
        failures.push(...localityStepFailures(step));
        const restored = applyDocumentOps(edit.document, edit.inverse);
        if (restored.isOk()) {
          const restoredParts = await serializedOpParts(restored.value.document);
          failures.push(
            ...serializedInverseStepFailures({ step, control, restored: restoredParts }),
          );
        }
        const edited = await serializedOpParts(edit.document);
        failures.push(
          ...serializedLocalityStepFailures({
            step,
            control,
            edited,
            documentPart: "word/document.xml",
          }),
        );
        if (failures.length > 0)
          throw new OperationPackageLawError({
            message: unexpectedFailure(kind, failures) ?? failures.join("\n"),
          });
      },
    );
    const fixture = fc.sample(packageDocumentArbitrary, { seed: 1336, numRuns: 1 }).at(0);
    const operationSeed = fc.sample(opSeedArbitrary, { seed: 1339, numRuns: 1 }).at(0);
    if (!fixture || !operationSeed)
      throw new OperationPackageLawError({ message: "Missing law fixture" });
    const examples = GENERATED_PACKAGE_OP_KINDS.flatMap((kind) =>
      GENERATED_PACKAGE_STORIES.map(
        (story): [typeof kind, typeof fixture, typeof operationSeed, typeof story] => [
          kind,
          fixture,
          operationSeed,
          story,
        ],
      ),
    );
    const expected = Object.entries(OPERATION_LAW_DISPOSITIONS).flatMap(([kind, disposition]) =>
      knownFailures(disposition, kind),
    );
    if (expected.length > 0)
      await assertKnownProperty(property, expected, { numRuns: examples.length + 135, examples });
    else await assertProperty(property, { numRuns: examples.length + 135, examples });
  },
  propertyTestTimeout(120_000),
);
