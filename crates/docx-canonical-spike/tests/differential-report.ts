/** Experimental comparison: every unported case is evidence, never a passing equivalence claim. */
import fc from "fast-check";
import { spawnSync } from "node:child_process";
import { appendFileSync, mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { isDeepStrictEqual } from "node:util";
import { applyDocumentOps } from "../../../packages/docx-core/src/ops/apply";
import { normalizeForOps } from "../../../packages/docx-core/src/ops/contract";
import { captureDocumentOp } from "../../../packages/docx-core/src/ops/wire";
import { parseDocx } from "../../../packages/core/src/docx/parser";
import { createDocx } from "../../../packages/core/src/docx/rezip";
import { operationPackageBytes } from "../../../scripts/lib/corpus-invariants/package-fixtures";
import { generateOpSequence } from "../../../scripts/lib/corpus-invariants/op-sequences";
import JSZip from "jszip";
import { Result } from "better-result";
import { EditorState } from "prosemirror-state";
import { CanonicalPublicOperations } from "../../../packages/core/src/controller/canonicalPublicOperations";
import {
  createCanonicalSession,
  publishCanonicalProjection,
} from "../../../packages/core/src/controller/canonicalSession";
import {
  createFolioAIEditSnapshot,
  createFolioAITextRangeHandle,
} from "../../../packages/core/src/ai-edits/snapshot";
import {
  BLOCK_COUNT,
  blockText,
  createBatchOverlapBatchArbitrary,
  token,
  type GeneratedOperation,
  type Span,
} from "../../../packages/core/src/__tests__/batchOverlapGenerators";
import type { FolioDocumentOperation } from "../../../packages/core/src/document-operations";
import type { FolioAIEditSnapshot } from "../../../packages/core/src/ai-edits/types";
import {
  documentArbitrary,
  opSeedArbitrary,
  opFor,
  trackedOpFor,
} from "../../../packages/docx-core/src/ops/__tests__/documentArbitraries";
import {
  GENERATED_PACKAGE_OP_KINDS,
  GENERATED_PACKAGE_STORIES,
  generatedCaseFor,
  packageDocumentArbitrary,
  captureDocumentArbitrary,
} from "../../../test/generators/packageOperationArbitraries";
import type { Document } from "../../../packages/docx-core/src/model/document";
import { encodeTagged, decodeTagged, withoutCaptureSymbols, transportScope } from "./harness-codec";
import type { DocumentOp } from "../../../packages/docx-core/src/ops/types";

const root = fileURLToPath(new URL("..", import.meta.url));
const repo = resolve(root, "../..");
const progress = process.env["RUST_SPIKE_PROGRESS_DIR"];
if (!progress) throw new TypeError("RUST_SPIKE_PROGRESS_DIR is required.");
mkdirSync(progress, { recursive: true });
const output = resolve(progress, "differential.jsonl");
writeFileSync(output, "");
const tally = new Map<string, number>();
const record = (value: unknown, classification: string) => {
  tally.set(classification, (tally.get(classification) ?? 0) + 1);
  appendFileSync(output, `${JSON.stringify(value)}\n`);
};
const native = (document: Document, ops: readonly DocumentOp[]): unknown => {
  const result = spawnSync(resolve(root, "target/debug/canonical-spike"), [], {
    input: `${JSON.stringify({ harness: { document: encodeTagged(document), ops: ops.map(captureDocumentOp) } })}\n`,
    encoding: "utf8",
  });
  if (result.error) throw result.error;
  if (result.status !== 0) throw new TypeError(`Native transport failed: ${result.stderr}`);
  const reply: unknown = JSON.parse(result.stdout);
  return typeof reply === "object" && reply !== null && "harness" in reply
    ? decodeTagged(reply.harness)
    : reply;
};
const differingPaths = (left: unknown, right: unknown, path = "$", out: string[] = []) => {
  if (isDeepStrictEqual(left, right)) return out;
  if (
    typeof left !== "object" ||
    left === null ||
    typeof right !== "object" ||
    right === null ||
    Array.isArray(left) !== Array.isArray(right)
  ) {
    out.push(path);
    return out;
  }
  for (const key of new Set([...Object.keys(left), ...Object.keys(right)])) {
    if (out.length >= 20) break;
    if (Object.hasOwn(left, key) !== Object.hasOwn(right, key)) out.push(`${path}.${key}:presence`);
    else differingPaths(Reflect.get(left, key), Reflect.get(right, key), `${path}.${key}`, out);
  }
  return out;
};
type CompareOptions = {
  family: string;
  seed: number;
  index: number;
  document: Document;
  ops: readonly DocumentOp[];
  phase?: string;
  source?: unknown;
  controlDocument?: Document;
};
const compare = ({
  family,
  seed,
  index,
  document,
  ops,
  phase = "forward",
  source,
  controlDocument,
}: CompareOptions) => {
  // The native request is measured before TS can mutate an input. Its scope
  // and complete request remain evidence even when the native arm refuses.
  const scope = transportScope(document);
  const request = {
    harness: { document: encodeTagged(document), ops: ops.map(captureDocumentOp) },
  };
  const actual = native(document, ops);
  const expected = applyDocumentOps(document, ops);
  const projected = expected.isErr()
    ? {
        status: "refused",
        opType: expected.error.opType,
        reason: expected.error.reason,
        message: expected.error.message,
      }
    : withoutCaptureSymbols(expected.value);
  const base = {
    family,
    seed,
    index,
    phase,
    opTypes: ops.map((op) => op.type),
    source,
    documentSource: phase === "inverse" ? "typescriptForward" : "generatedInput",
    ...scope,
  };
  if (
    typeof actual === "object" &&
    actual !== null &&
    "status" in actual &&
    actual.status === "unsupported"
  ) {
    record(
      {
        ...base,
        classification: "unsupported",
        actual,
        expected: encodeTagged(projected),
        request,
      },
      `unsupported:${"dimension" in actual ? actual.dimension : "unknown"}`,
    );
  } else if (!isDeepStrictEqual(actual, projected)) {
    record(
      {
        ...base,
        classification: "mismatch",
        paths: differingPaths(encodeTagged(actual), encodeTagged(projected)),
        actual: encodeTagged(actual),
        expected: encodeTagged(projected),
        request,
      },
      `mismatch:${ops.at(0)?.type ?? "batch"}`,
    );
  } else {
    record({ ...base, classification: "equal" }, "equal");
  }
  if (
    expected.isOk() &&
    controlDocument !== undefined &&
    !isDeepStrictEqual(
      withoutCaptureSymbols(expected.value.document),
      withoutCaptureSymbols(controlDocument),
    )
  ) {
    record(
      {
        ...base,
        classification: "typescript-control-mismatch",
        request,
        actualDocument: encodeTagged(expected.value.document),
        controlDocument: encodeTagged(controlDocument),
        paths: differingPaths(encodeTagged(expected.value.document), encodeTagged(controlDocument)),
      },
      "typescript-control-mismatch",
    );
  }
  if (expected.isOk() && phase === "forward")
    compare({
      family,
      seed,
      index,
      document: expected.value.document,
      ops: expected.value.inverse,
      phase: "inverse",
      source,
      controlDocument: document,
    });
};

type PinnedReplay = { seed: number; path?: string; title: string; run: "pinned" | "generated" };
const replaysFor = (relative: string, title: string): PinnedReplay[] => {
  const registry: unknown = JSON.parse(readFileSync(resolve(repo, relative), "utf8"));
  if (typeof registry !== "object" || registry === null || Array.isArray(registry))
    throw new TypeError("Invalid property seed registry.");
  const entries: unknown = Reflect.get(registry, title) ?? [];
  if (!Array.isArray(entries)) throw new TypeError(`Invalid seed list for ${title}.`);
  const pinned = entries.map((entry: unknown): PinnedReplay => {
    if (
      typeof entry !== "object" ||
      entry === null ||
      !("seed" in entry) ||
      typeof entry.seed !== "number" ||
      !Number.isInteger(entry.seed)
    )
      throw new TypeError(`Invalid pinned seed for ${title}.`);
    const path = "path" in entry ? entry.path : undefined;
    if (path !== undefined && typeof path !== "string") throw new TypeError("Invalid shrink path.");
    if (path === undefined) return { seed: entry.seed, title, run: "pinned" };
    return { seed: entry.seed, path, title, run: "pinned" };
  });
  return [...pinned, { seed: 20261005, title, run: "generated" }];
};
type SampleReplayOptions<Ts> = {
  property: fc.IRawProperty<Ts>;
  replay: PinnedReplay;
  numRuns: number;
};
const numRunsFactor = Number(process.env["PROPERTY_TEST_NUM_RUNS_FACTOR"] ?? 1);
if (!Number.isFinite(numRunsFactor) || numRunsFactor < 1)
  throw new TypeError("Property run factor must be finite and at least one.");
const sampleReplay = <Ts>({ property, replay, numRuns }: SampleReplayOptions<Ts>) => {
  const sampled = Result.try(() =>
    fc.sample(property, {
      seed: replay.seed,
      ...(replay.path === undefined ? {} : { path: replay.path }),
      numRuns: Math.ceil(numRuns * numRunsFactor),
    }),
  );
  if (sampled.isOk()) return sampled.value;
  record(
    {
      family: "generator-replay",
      classification: "gap",
      source: replay,
      message: String(sampled.error),
      reason: "Pinned shrink path cannot be sampled in the current source domain.",
    },
    "gap:generatorReplay",
  );
  return [];
};
const seedsIn = (relative: string): number[] => {
  const data: unknown = JSON.parse(readFileSync(resolve(repo, relative), "utf8"));
  if (typeof data !== "object" || data === null) throw new TypeError("Invalid seed registry.");
  const seeds = new Set<number>();
  for (const records of Object.values(data)) {
    if (!Array.isArray(records)) throw new TypeError("Invalid seed list.");
    for (const entry of records) if (typeof entry?.seed === "number") seeds.add(entry.seed);
  }
  return [...seeds];
};
const applySeeds = [
  20261005,
  ...seedsIn(
    "test/property-seeds/packages%2Fdocx-core%2Fsrc%2Fops%2F__tests__%2Fapply.property.test.ts.json",
  ),
];
for (const seed of applySeeds) {
  const draws = fc.sample(fc.tuple(documentArbitrary, opSeedArbitrary), { seed, numRuns: 80 });
  draws.forEach(([document, generated], index) => {
    compare({ family: "apply.property", seed, index, document, ops: [opFor(document, generated)] });
    compare({
      family: "review.property",
      seed,
      index,
      document,
      ops: [trackedOpFor(document, generated, index)],
    });
  });
  // Sequential generation uses each TS state, never regenerates or changes the operation to fit Rust.
  let current = draws.at(0)?.[0];
  if (!current || !("package" in current))
    throw new TypeError("Missing generated sequence document.");
  const original = current;
  const composed: DocumentOp[] = [];
  const refusals: string[] = [];
  for (const [index, [, generated]] of draws.slice(0, 12).entries()) {
    const op = opFor(current, generated);
    compare({ family: "apply.sequence", seed, index, document: current, ops: [op] });
    const applied = applyDocumentOps(current, [op]);
    if (applied.isOk()) {
      composed.push(op);
      current = applied.value.document;
    } else refusals.push(`${op.type}:${applied.error.reason}`);
  }
  compare({
    family: "apply.sequence.atomic",
    seed,
    index: 0,
    document: original,
    ops: composed,
    controlDocument: current,
    source: { generator: "opFor on each TS state", successfulSteps: composed.length, refusals },
  });
}

const packageRegistry = "test/property-seeds/scripts%2Fpackage-operation-laws.test.ts.json";
const packageReplays: PinnedReplay[] = [];
type ComparePackagedOptions = {
  args: Parameters<typeof generatedCaseFor>[0];
  replay: PinnedReplay;
  index: number;
};
const packageGeneratorOrder = (title: string) => {
  switch (title) {
    case "every operation preserves package inverse and declared scope laws":
      return ["kind", "document", "operationSeed", "story"];
    case "created story removal restores authored package registrations exactly":
      return ["document", "operationSeed"];
    case "every declared operation and story pair preserves package laws":
      return [
        "independent document sample",
        "independent operationSeed sample",
        "kind/story census",
      ];
    default:
      return ["document", "operationSeed", "story"];
  }
};
const comparePackaged = async ({ args, replay, index }: ComparePackagedOptions) => {
  // The package-law suite generates authored ZIP bytes before selecting an op.
  // Sampling the same arbitrary without this parse loses its captured inputs.
  const bytes = await operationPackageBytes(args.document, args.seed);
  const parsed = normalizeForOps(await parseDocx(bytes, { preloadFonts: false }));
  const generated = generatedCaseFor({ ...args, document: parsed });
  compare({
    family: "package-operation-laws",
    seed: replay.seed,
    index,
    document: generated.document,
    ops: [generated.op],
    source: {
      ...replay,
      generatorOrder: packageGeneratorOrder(replay.title),
      pipeline: "operationPackageBytes -> parseDocx -> normalizeForOps -> generatedCaseFor",
      story: args.story,
      ...(replay.title === "every declared operation and story pair preserves package laws"
        ? { independentDocumentSeed: 1336, independentOperationSeed: 1339 }
        : {}),
    },
  });
};

// fc.sample accepts the source property's seed/path and walks its shrink tree.
// Predicates are not executed: these are differential inputs, not law verdicts.
const allPackageProperty = fc.asyncProperty(
  fc.constantFrom(...GENERATED_PACKAGE_OP_KINDS),
  packageDocumentArbitrary,
  opSeedArbitrary,
  fc.constantFrom(...GENERATED_PACKAGE_STORIES),
  async () => {},
);
for (const replay of replaysFor(
  packageRegistry,
  "every operation preserves package inverse and declared scope laws",
)) {
  packageReplays.push(replay);
  for (const [index, [kind, document, seed, story]] of sampleReplay({
    property: allPackageProperty,
    replay,
    numRuns: 135,
  }).entries())
    await comparePackaged({ args: { kind, document, seed, story }, replay, index });
}
for (const [kind, title] of [
  ["setParagraphProps", "paragraph property package inverses retain authored stories"],
  ["joinBlocks", "join package inverses retain authored stories"],
] as const) {
  const property = fc.asyncProperty(
    captureDocumentArbitrary,
    opSeedArbitrary,
    fc.constantFrom(...GENERATED_PACKAGE_STORIES),
    async () => {},
  );
  for (const replay of replaysFor(packageRegistry, title)) {
    packageReplays.push(replay);
    for (const [index, [document, seed, story]] of sampleReplay({
      property,
      replay,
      numRuns: 100,
    }).entries())
      await comparePackaged({ args: { kind, document, seed, story }, replay, index });
  }
}
const removalProperty = fc.asyncProperty(packageDocumentArbitrary, opSeedArbitrary, async () => {});
for (const replay of replaysFor(
  packageRegistry,
  "created story removal restores authored package registrations exactly",
)) {
  packageReplays.push(replay);
  for (const [index, [document, seed]] of sampleReplay({
    property: removalProperty,
    replay,
    numRuns: 30,
  }).entries())
    await comparePackaged({
      args: { kind: "removeHeaderFooter", document, seed, story: "main" },
      replay,
      index,
    });
}
// The explicit pair census uses two independent samples, as its source test does.
const pairDocument = fc.sample(packageDocumentArbitrary, { seed: 1336, numRuns: 1 }).at(0);
const pairSeed = fc.sample(opSeedArbitrary, { seed: 1339, numRuns: 1 }).at(0);
if (!pairDocument || !pairSeed) throw new TypeError("Missing explicit operation/story fixture.");
for (const [kindIndex, kind] of GENERATED_PACKAGE_OP_KINDS.entries())
  for (const [storyIndex, story] of GENERATED_PACKAGE_STORIES.entries())
    await comparePackaged({
      args: { kind, document: pairDocument, seed: pairSeed, story },
      replay: {
        seed: 1336,
        title: "every declared operation and story pair preserves package laws",
        run: "generated",
      },
      index: kindIndex * GENERATED_PACKAGE_STORIES.length + storyIndex,
    });

const packageSequenceProperty = fc.asyncProperty(
  captureDocumentArbitrary,
  fc.integer({ min: 0, max: 0x7fffffff }),
  fc.constantFrom(" ", "\n", "\r\n", "\t"),
  async () => {},
);
for (const replay of replaysFor(
  packageRegistry,
  "operation sequence generation retains the parsed package control",
)) {
  packageReplays.push(replay);
  for (const [index, [document, seed, whitespace]] of sampleReplay({
    property: packageSequenceProperty,
    replay,
    numRuns: 12,
  }).entries()) {
    const zip = await JSZip.loadAsync(await createDocx(document));
    const xml = await zip.file("word/document.xml")?.async("text");
    if (!xml) throw new TypeError("Missing sequence document XML.");
    zip.file("word/document.xml", xml.replaceAll("<w:pPr>", `<w:pPr>${whitespace}`));
    const bytes = await zip.generateAsync({ type: "arraybuffer" });
    const parsed = normalizeForOps(await parseDocx(bytes, { preloadFonts: false }));
    const sequence = generateOpSequence(parsed, seed);
    const ops = sequence.steps.map(({ op }) => op);
    compare({
      family: "package-operation-laws.sequence.atomic",
      seed: replay.seed,
      index,
      document: sequence.original,
      ops,
      controlDocument: sequence.document,
      source: {
        ...replay,
        sequenceSeed: seed,
        whitespace,
        generator: "generateOpSequence",
        successfulSteps: ops.length,
        mutations: sequence.mutations,
        refusals: sequence.refusals,
      },
    });
    // Retain the generator's own composed inverse as a separate source arm;
    // forward comparison already checks the atomic TS batch's inverse.
    compare({
      family: "package-operation-laws.sequence.composed-inverse",
      seed: replay.seed,
      index,
      document: sequence.document,
      ops: sequence.inverse,
      phase: "inverse",
      controlDocument: sequence.original,
      source: { ...replay, sequenceSeed: seed, generator: "generateOpSequence.inverse" },
    });
  }
}

const spanText = ({ block, first, last }: Span) =>
  Array.from({ length: last - first + 1 }, (_, offset) => token(block, first + offset)).join(" ");
type MaterializeBatchOptions = {
  snapshot: FolioAIEditSnapshot;
  blockIds: readonly string[];
  generated: GeneratedOperation;
  index: number;
};
/** Same public recipes as batchOverlap.test.ts; no replacement primitive is invented here. */
const materializeBatchOperation = ({
  snapshot,
  blockIds,
  generated,
  index,
}: MaterializeBatchOptions): FolioDocumentOperation | null => {
  const id = `op${String(index)}`;
  const blockId = blockIds[generated.block] ?? "";
  const current = snapshot.blocks.find((block) => block.id === blockId);
  if (!current) return null;
  const rangeOf = (span: Span) => {
    const text = spanText(span);
    const start = current.text.indexOf(text);
    return start < 0
      ? null
      : createFolioAITextRangeHandle({
          blockId,
          text: current.text,
          startOffset: start,
          endOffset: start + text.length,
        });
  };
  switch (generated.kind) {
    case "replaceInBlock": {
      const find = spanText(generated);
      return current.text.includes(find)
        ? { id, type: "replaceInBlock", blockId, find, replace: `new${String(index)}` }
        : null;
    }
    case "replaceRange": {
      const range = rangeOf(generated);
      return range && { id, type: "replaceRange", range, replace: `rng${String(index)}` };
    }
    case "formatRange": {
      const range = rangeOf(generated);
      return range && { id, type: "formatRange", range, formatting: { bold: true } };
    }
    case "commentOnRange": {
      const range = rangeOf(generated);
      return range && { id, type: "commentOnRange", range, comment: { text: `c${String(index)}` } };
    }
    case "commentOnBlock":
      return { id, type: "commentOnBlock", blockId, comment: { text: `c${String(index)}` } };
    case "splitBlock": {
      const previous = token(generated.block, generated.before - 1);
      const at = current.text.indexOf(previous);
      return at < 0
        ? null
        : { id, type: "splitBlock", blockId, offset: at + previous.length, separator: " " };
    }
    case "mergeBlockWithNext":
      return {
        id,
        type: generated.kind,
        blockId,
        ...(generated.separator === undefined ? {} : { separator: generated.separator }),
      };
    case "deleteBlock":
      return { id, type: generated.kind, blockId };
    case "replaceBlock":
      return { id, type: "replaceBlock", blockId, text: `whole${String(index)}` };
    case "setBlockParagraphProperties":
      return { id, type: generated.kind, blockId, properties: { alignment: "center" } };
    case "insertAfterBlock":
    case "insertBeforeBlock":
      return { id, type: generated.kind, blockId, text: `ins${String(index)}` };
  }
};

const batchRegistry =
  "test/property-seeds/packages%2Fcore%2Fsrc%2Fai-edits%2FbatchOverlap.test.ts.json";
const batchReplays = replaysFor(
  batchRegistry,
  "refuses each conflict and applies the rest as one at a time would",
);
const batchProperty = fc.asyncProperty(
  createBatchOverlapBatchArbitrary(),
  fc.constantFrom("direct" as const, "tracked-changes" as const),
  async () => {},
);
const batchDocument = normalizeForOps(
  await parseDocx(
    await createDocx({
      package: {
        document: {
          content: Array.from({ length: BLOCK_COUNT }, (_, block) => ({
            type: "paragraph" as const,
            paraId: (0x1000_0000 + block).toString(16).toUpperCase(),
            content: [
              {
                type: "run" as const,
                content: [{ type: "text" as const, text: blockText(block) }],
              },
            ],
          })),
        },
      },
    }),
    { preloadFonts: false },
  ),
);
record(
  {
    family: "batchOverlap",
    classification: "gap",
    gap: "publicOps.headlessSession",
    source: "packages/core/src/__tests__/operationBatchDocuments.ts:OperationSession.apply",
    pipeline:
      "applyFolioDocumentOperations -> applyFolioAIEditOperations -> ProseMirror transactions",
    message:
      "The source batch-overlap suite does not emit canonical operations. The compiler arm below is a separate canonical public-batch path over the same sampled recipes; it cannot establish equivalence of the source PM overlap oracle.",
  },
  "gap:publicOps.headlessSession",
);
for (const replay of batchReplays) {
  for (const [index, [generated, mode]] of sampleReplay({
    property: batchProperty,
    replay,
    numRuns: 150,
  }).entries()) {
    const created = createCanonicalSession(batchDocument);
    if (created.isErr()) {
      record(
        {
          family: "batchOverlap.canonical-compiler",
          classification: "gap",
          source: replay,
          index,
          generated,
          mode,
          error: created.error.message,
        },
        "gap:canonicalBatchSession",
      );
      continue;
    }
    const session = created.value;
    const original = session.document;
    let state = EditorState.create({ doc: session.projection.doc });
    let compiledOps: readonly DocumentOp[] | undefined;
    // Transparent test instrumentation observes the real compiler handoff;
    // its resolver, overlap checks, ordering, allocator and primitives run unchanged.
    const prepareOps = session.prepareOps.bind(session);
    session.prepareOps = (currentState, ops, selection) => {
      compiledOps = ops;
      return prepareOps(currentState, ops, selection);
    };
    const executor = new CanonicalPublicOperations({
      session,
      getState: () => state,
      publish: (commit) => {
        const published = publishCanonicalProjection({ state, commit, session });
        if (published.isErr()) return false;
        state = published.value.state;
        return true;
      },
    });
    const snapshot = createFolioAIEditSnapshot(state.doc);
    const blockIds = snapshot.blocks.map((block) => block.id);
    const operations = generated.flatMap((operation, operationIndex) => {
      const materialized = materializeBatchOperation({
        snapshot,
        blockIds,
        generated: operation,
        index: operationIndex,
      });
      return materialized === null ? [] : [materialized];
    });
    if (operations.length !== generated.length) {
      record(
        {
          family: "batchOverlap.canonical-compiler",
          classification: "gap",
          source: replay,
          index,
          generated,
          mode,
          operations,
          reason: "A source recipe did not materialize against its token fixture.",
        },
        "gap:batchMaterialization",
      );
      continue;
    }
    const result = executor.apply({
      snapshot,
      batch: { version: 1, mode, operations },
      author: "Reviewer",
      revisionStamp: { date: "2026-01-01T00:00:00Z", idSeed: 1 },
    });
    const source = {
      ...replay,
      generated,
      mode,
      publicOperations: operations,
      publicResult: result,
      pipeline: "CanonicalPublicOperations -> compileEditorIntent -> CanonicalSession.prepareOps",
      originalPipelineGap: "publicOps.headlessSession",
    };
    if (result.skipped.length > 0)
      record(
        {
          family: "batchOverlap.canonical-compiler",
          classification: "compiler-skips",
          index,
          source,
        },
        "compiler-skips",
      );
    if (compiledOps === undefined || result.applied.length === 0) {
      record(
        {
          family: "batchOverlap.canonical-compiler",
          classification: "gap",
          index,
          source,
          reason: "No committed primitive batch reached the canonical journal.",
        },
        "gap:batchCanonicalOps",
      );
      continue;
    }
    compare({
      family: "batchOverlap.canonical-compiler.atomic",
      seed: replay.seed,
      index,
      document: original,
      ops: compiledOps,
      source,
      controlDocument: session.document,
    });
  }
}
writeFileSync(
  resolve(progress, "differential-summary.json"),
  `${JSON.stringify(
    {
      time: `${new Date().toLocaleString("sv-SE", { timeZone: "Europe/Prague" })} CEST`,
      tally: Object.fromEntries(tally),
      applySeeds,
      packageReplays,
      batchReplays,
      numRunsFactor,
      note: "Harness codec preserves Map/Date/binary/ownedundefined. Capture symbols excluded and counted; shared references counted. Every native mismatch and unsupported result remains evidence, with its request and TS result. Inverses are compared independently from the TS forward state even when the native forward fails. Package draws use the source property's arbitrary order, pinned seed/path, authored package parse and run count. Batch draws use the shared source arbitrary and pinned seed/path; the original PM/headless pipeline gap is explicit. Compiler-emitted primitive batches are a separate arm, not equivalence of the PM overlap suite. These model comparisons make no serialized-package-byte law claim.",
    },
    null,
    2,
  )}\n`,
);
console.log(Object.fromEntries(tally));
