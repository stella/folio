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
import { cloneDocumentWithParagraphPropertySources } from "@stll/folio-core/docx/document-clone";
import { visitDocxParagraphs } from "../packages/core/src/docx/paragraphTraversal";
import { applyDocumentOps } from "../packages/docx-core/src/ops/apply";
import { DOCUMENT_OP_TYPES } from "../packages/docx-core/src/ops/types";
import { failureMarker } from "../test/consumer-scenarios/support/failure-fingerprints";
import { opSeedArbitrary } from "../packages/docx-core/src/ops/__tests__/documentArbitraries";
import {
  generatedCaseFor,
  GENERATED_PACKAGE_OP_KINDS,
  GENERATED_PACKAGE_STORIES,
  packageDocumentArbitrary,
  captureDocumentArbitrary,
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
  if (kind === "joinBlocks" || kind === "setParagraphProps") {
    // Anchor the inverse control to authored bytes, so losing captures before
    // the baseline cannot make both sides agree on an already damaged package.
    const source = await JSZip.loadAsync(bytes);
    for (const part of source.file(/^word\/(?:document|header[^/]*|footer[^/]*)\.xml$/u)) {
      const xml = await part.async("text");
      // The XML parser normalizes CR and CRLF before capturing paragraph XML.
      const capturedPrefix = xml
        .replace(/\r\n?/gu, "\n")
        .match(/<w:pPr>\s+/u)
        ?.at(0);
      if (capturedPrefix === undefined)
        throw new OperationPackageLawError({
          message: "Authored paragraph whitespace is missing.",
        });
      const saved = control.get(part.name);
      if (saved === undefined)
        throw new OperationPackageLawError({ message: `The baseline lost ${part.name}.` });
      expect(new TextDecoder().decode(saved).replace(/\r\n?/gu, "\n")).toContain(capturedPrefix);
    }
  }
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
  "every declared operation and story pair preserves package laws",
  async () => {
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
    const exercised = new Set<string>();
    for (const [kind, document, seed, story] of examples) {
      await assertPackageOperationLaws({ kind, document, seed, story });
      exercised.add(JSON.stringify([kind, story]));
    }
    expect(exercised).toEqual(
      new Set(
        GENERATED_PACKAGE_OP_KINDS.flatMap((kind) =>
          GENERATED_PACKAGE_STORIES.map((story) => JSON.stringify([kind, story])),
        ),
      ),
    );
    expect(exercised.size).toBeGreaterThan(0);
  },
  propertyTestTimeout(120_000),
);

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
    const expected = Object.entries(OPERATION_LAW_DISPOSITIONS).flatMap(([kind, disposition]) =>
      knownFailures(disposition, kind),
    );
    if (expected.length > 0) await assertKnownProperty(property, expected, { numRuns: 135 });
    else await assertProperty(property, { numRuns: 135 });
  },
  propertyTestTimeout(120_000),
);

test(
  "paragraph property package inverses retain authored stories",
  async () => {
    let cases = 0;
    await assertProperty(
      fc.asyncProperty(
        captureDocumentArbitrary,
        opSeedArbitrary,
        fc.constantFrom(...GENERATED_PACKAGE_STORIES),
        async (document, seed, story) => {
          cases += 1;
          await assertPackageOperationLaws({ kind: "setParagraphProps", document, seed, story });
        },
      ),
      { numRuns: 100 },
    );
    expect(cases).toBeGreaterThan(0);
  },
  propertyTestTimeout(60_000),
);

test(
  "join package inverses retain authored stories",
  async () => {
    let cases = 0;
    await assertProperty(
      fc.asyncProperty(
        captureDocumentArbitrary,
        opSeedArbitrary,
        fc.constantFrom(...GENERATED_PACKAGE_STORIES),
        async (document, seed, story) => {
          cases += 1;
          await assertPackageOperationLaws({ kind: "joinBlocks", document, seed, story });
        },
      ),
      { numRuns: 100 },
    );
    expect(cases).toBeGreaterThan(0);
  },
  propertyTestTimeout(60_000),
);

test(
  "operation sequence generation retains the parsed package control",
  async () => {
    let cases = 0;
    await assertProperty(
      fc.asyncProperty(
        captureDocumentArbitrary,
        fc.integer({ min: 0, max: 0x7fffffff }),
        fc.constantFrom(" ", "\n", "\r\n", "\t"),
        async (document, seed, whitespace) => {
          cases += 1;
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
          const withoutCaptures = cloneDocumentWithParagraphPropertySources(sequence.original);
          let removed = 0;
          visitDocxParagraphs({ documentBody: withoutCaptures.package.document }, (paragraph) => {
            for (const key of Object.getOwnPropertySymbols(paragraph)) {
              if (key.description !== "paragraphPropertyCapture") continue;
              expect(Reflect.deleteProperty(paragraph, key)).toBe(true);
              removed++;
            }
          });
          expect(removed).toBeGreaterThan(0);
          expect(structuredClone(withoutCaptures.package.document.content)).toStrictEqual(
            structuredClone(sequence.original.package.document.content),
          );
          const mutated = await serializedOpParts(withoutCaptures);
          expect(mutated.get("word/document.xml")).not.toEqual(control.get("word/document.xml"));
        },
      ),
      { numRuns: 12 },
    );
    expect(cases).toBeGreaterThan(0);
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
