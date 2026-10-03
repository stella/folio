/** Package byte and scope laws extend the corpus oracle to every operation kind in unit CI. */
import { expect, test } from "bun:test";
import fc from "fast-check";
import JSZip from "jszip";
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
  generateOpSequence,
  exactOpModel,
  sameOpModel,
  serializeOpDocument,
  serializedOpParts,
} from "./lib/corpus-invariants/op-sequences";
import { operationPackageBytes } from "./lib/corpus-invariants/package-fixtures";
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
} from "../test/generators/packageOperationArbitraries";

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

type UnexpectedFailureOptions = {
  kind: keyof typeof OPERATION_LAW_DISPOSITIONS;
  failures: readonly string[];
  disposition?: OperationLawDisposition;
};
const unexpectedFailure = ({
  kind,
  failures,
  disposition = OPERATION_LAW_DISPOSITIONS[kind],
}: UnexpectedFailureOptions) => {
  const expected = knownFailures(disposition, kind);
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
  expect(
    unexpectedFailure({
      kind: "joinBlocks",
      failures: [inverse, scope],
      disposition: { knownIssue: "T4", fingerprint: "af2bdcaabf37424e" },
    }),
  ).toBe(scope);
  expect(unexpectedFailure({ kind: "insertText", failures: [inverse] })).toBe(inverse);
});

const assertPackageOperationLaws = async ({
  kind,
  document,
  seed,
  story,
}: Parameters<typeof generatedCaseFor>[0]) => {
  const bytes = await operationPackageBytes(document, seed);
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
    failures.push(...serializedInverseStepFailures({ step, control, restored: restoredParts }));
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
      message: unexpectedFailure({ kind, failures }) ?? failures.join("\n"),
    });
};

test(
  "every operation preserves package inverse and declared scope laws",
  async () => {
    expect(new Set(GENERATED_PACKAGE_OP_KINDS)).toEqual(new Set(Object.values(DOCUMENT_OP_TYPES)));
    const property = fc.asyncProperty(
      fc.constantFrom(...GENERATED_PACKAGE_OP_KINDS),
      packageDocumentArbitrary,
      opSeedArbitrary,
      fc.constantFrom(...GENERATED_PACKAGE_STORIES),
      async (kind, document, seed, story) =>
        assertPackageOperationLaws({ kind, document, seed, story }),
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

test(
  "operation sequence generation retains the parsed package control",
  async () => {
    await assertProperty(
      fc.asyncProperty(
        packageDocumentArbitrary,
        fc.integer({ min: 0, max: 0x7fffffff }),
        fc.constantFrom(" ", "\n", "\r\n", "\t"),
        async (document, seed, whitespace) => {
          const zip = await JSZip.loadAsync(await createDocx(document));
          const xml = await zip.file("word/document.xml")?.async("text");
          if (!xml) throw new TypeError("Missing document part");
          expect(xml).toContain("<w:pPr>");
          zip.file("word/document.xml", xml.replaceAll("<w:pPr>", `<w:pPr>${whitespace}`));
          const bytes = await zip.generateAsync({ type: "arraybuffer" });
          const parsed = normalizeForOps(await parseDocx(bytes, { preloadFonts: false }));
          const control = await serializedOpParts(parsed);
          const sequence = generateOpSequence(parsed, seed);
          const initial = await serializedOpParts(sequence.original);
          expect(initial).toEqual(control);
        },
      ),
      { numRuns: 12 },
    );
  },
  propertyTestTimeout(60_000),
);

test(
  "created story removal restores authored package registrations exactly",
  async () => {
    await assertProperty(
      fc.asyncProperty(packageDocumentArbitrary, opSeedArbitrary, async (document, seed) => {
        await assertPackageOperationLaws({
          kind: "removeHeaderFooter",
          document,
          seed,
          story: "main",
        });
      }),
      { numRuns: 30 },
    );
  },
  propertyTestTimeout(30_000),
);
