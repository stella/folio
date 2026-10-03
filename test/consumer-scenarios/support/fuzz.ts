/**
 * The seeded random flows the fuzz scenarios run: chains of public
 * operations (tool calls, core batches, model mistakes, comments,
 * accept/reject, save and reopen) over the synthetic documents, checking
 * after every step that the document saves, reopens to what the reviewer
 * showed, and reads alike everywhere (`STEP_CHECKS`), and that every applied
 * operation did what it asked (support/oracle.ts). A flow is fully
 * determined by its seed, step count, kind and generation: a `"collisions"`
 * flow draws half its batches from `COLLISIONS`.
 *
 * The `"targeted"` generation (the default) aims most operations at edges
 * (support/targets.ts): recently touched blocks, story edges, section
 * breaks, table edges, pending revisions, comment anchors, inline objects,
 * fields and list boundaries, at run, note-reference and surrogate
 * boundaries. It also edits headers, footers and notes, and changes session
 * mid-flow: a save reopened by a new reviewer, and a selective save the same
 * reviewer keeps working after. The `"legacy"` generation draws exactly what
 * flows drew before targeting, so seeds pinned then replay unchanged.
 *
 * Every run also records itself as a flow file (support/flow-file.ts):
 * `runFlowFile` replays one, which is what shrinking, the corpus and the
 * checked-in flows run.
 */

import { createReviewerBridge, executeFolioToolCallUntyped } from "@stll/folio-agents";
import type { FolioDocumentStoryHandle } from "@stll/folio-core/server";
import { readFileSync } from "node:fs";

import { observeHits, registerFeatureOperations, type StepKind } from "./coverage.ts";
import {
  featureCellKey,
  featureOperationHits,
  parseFeatureCoverage,
  gapWeight,
  generatedSelection,
  targetFeatureSignature,
  type FeatureCoverage,
} from "./feature-coverage.ts";
import { COLLISION_FIXTURES, FIXTURES, openReviewer, STORY_FIXTURES } from "./documents.ts";
import { normalizeAssertion } from "./failure-fingerprints.ts";
import {
  type Action,
  type FlowFile,
  type FlowKind,
  type FlowStep,
  type Generation,
  LEGACY_ACTIONS,
  TARGETED_ACTIONS,
} from "./flow-file.ts";
import { assertReadersAgree, saveAndReopen, visibleState } from "./invariants.ts";
import { LARGE_DOCUMENT_FIXTURE, largeDocument } from "./large-document.ts";
import { startRelations } from "./metamorphic.ts";
import { assertRequestedOutcome, assertResolvedTo, capture, captureResolution } from "./oracle.ts";
import {
  type Block,
  COLLISIONS,
  GENERATORS,
  coreBatch,
  MISTAKES,
  MODES,
  type Mode,
  type Operation,
  randomOperation,
  supports,
} from "./operations.ts";
import { createRandom, type Random, sentence } from "./random.ts";
import {
  isPublicCorpusFixture,
  loadPublicCorpusFixture,
  type PublicCorpusFixture,
} from "./public-corpus.ts";
import { drawSwarm, swarmIncludesBatch } from "./swarm.ts";
import { biasedPicker, blocksOfStory, featureIndex, storyKindOf, type Picker } from "./targets.ts";

type Reviewer = Awaited<ReturnType<typeof openReviewer>>;

export type { FlowKind, Generation } from "./flow-file.ts";

export type Flow = {
  reviewer: Reviewer;
  mode: Mode;
  random: Random;
  /** Ids read earlier that may no longer name a block. */
  seenIds: Set<string>;
  /** What each step did, written before it runs so a throw still names it. */
  log: string[];
  kind: FlowKind;
  generation: Generation;
  /** Blocks earlier steps aimed at, most recent last. */
  recent: string[];
  /** The session the next step runs in. */
  session: StepKind;
  /** The steps so far, as a flow file records them (the running one last). */
  trace: FlowStep[];
  /** The running step, when a flow file planned it. */
  planned: FlowStep | undefined;
  /** Ids of the fixture's blocks, which every run of the flow opens with. */
  fixtureIds: ReadonlySet<string>;
  swarm: string[] | undefined;
  weights: FeatureCoverage | undefined;
  operationHits: Record<string, number> | undefined;
};

const MAIN: FolioDocumentStoryHandle = { type: "main" };

/** Read weights only for fresh seeded flows, never while replaying a recorded flow. */
const processFeatureWeights = (): FeatureCoverage | undefined => {
  const weightsPath = process.env["FOLIO_SCENARIO_FEATURE_WEIGHTS"];
  return weightsPath
    ? parseFeatureCoverage(JSON.parse(readFileSync(weightsPath, "utf8")))
    : undefined;
};

const blocksOf = (flow: Flow): Block[] => flow.reviewer.getContent() as Block[];

/** The picker a targeted flow aims with, over `story`'s blocks as they are now. */
const pickerFor = (
  flow: Flow,
  story: FolioDocumentStoryHandle = MAIN,
): Picker | ((type: string) => Picker) | undefined => {
  if (flow.generation === "legacy") return undefined;
  const index = featureIndex(flow.reviewer, story);
  const options = { index, recent: flow.recent };
  const coverageWeights = flow.weights;
  if (!coverageWeights) return biasedPicker(flow.random, options);
  return (operation) =>
    biasedPicker(flow.random, {
      ...options,
      coverageWeight: (block) => {
        const features = targetFeatureSignature(
          block,
          index.features.get(block.id) ?? new Set(),
          storyKindOf(index, block),
        );
        return Math.max(
          ...features.map((feature) =>
            gapWeight(
              coverageWeights.cells[
                featureCellKey({
                  operation,
                  feature,
                  selection: generatedSelection(operation),
                })
              ] ?? 0,
            ),
          ),
        );
      },
    });
};

/** Remember the blocks `operations` aimed at, for the next steps to aim near. */
const touch = (flow: Flow, operations: readonly Operation[]): void => {
  for (const operation of operations) {
    const range = operation["range"] as { blockId?: unknown } | undefined;
    const id = operation["blockId"] ?? range?.blockId;
    if (typeof id === "string") flow.recent.push(id);
  }
  flow.recent.splice(0, Math.max(0, flow.recent.length - 12));
};

/**
 * The batch a step applies: what it draws or, when its flow file pins
 * operations, those (drawn anyway, so the step's later draws stay put).
 * Recorded on the step's trace entry either way.
 */
const randomOperations = (
  flow: Flow,
  blocks: readonly Block[] = blocksOf(flow),
  pick: Picker | ((type: string) => Picker) | undefined = pickerFor(flow),
): Operation[] => {
  const drawn = drawOperations(flow, blocks, pick);
  const pinned = flow.planned?.operations;
  const used = pinned === undefined ? drawn : (fromPositions(pinned, blocks) as Operation[]);
  if (!swarmIncludesBatch(flow.swarm, used)) {
    throw new TypeError("flow file: pinned operation is disabled by its swarm");
  }
  const traced = flow.trace.at(-1);
  if (traced !== undefined) {
    traced.operations = toPositions(used, blocks, flow.fixtureIds) as Operation[];
  }
  return used;
};

// A block a flow inserted gets a new id on every run, so a pinned operation
// names it by its position in the story's blocks when the step drew, `@<n>`.
// A fixture block keeps its id, which survives steps before it going away.
const POSITION = /^@(\d+)$/u;

const mapBlockIds = (value: unknown, map: (id: string) => string): unknown => {
  if (Array.isArray(value)) return value.map((item) => mapBlockIds(item, map));
  if (typeof value !== "object" || value === null) return value;
  return Object.fromEntries(
    Object.entries(value).map(([key, item]) => [
      key,
      key === "blockId" && typeof item === "string" ? map(item) : mapBlockIds(item, map),
    ]),
  );
};

const toPositions = (
  operations: readonly Operation[],
  blocks: readonly Block[],
  fixtureIds: ReadonlySet<string>,
): unknown =>
  mapBlockIds(operations, (id) => {
    if (fixtureIds.has(id)) return id;
    const index = blocks.findIndex((block) => block.id === id);
    return index === -1 ? id : `@${index}`;
  });

const fromPositions = (operations: readonly unknown[], blocks: readonly Block[]): unknown =>
  mapBlockIds(operations, (id) => {
    const index = POSITION.exec(id)?.[1];
    return index === undefined ? id : (blocks[Number(index)]?.id ?? id);
  });

/** Every block id of every story the reviewer opened with. */
const blockIdsOf = (reviewer: Reviewer): Set<string> =>
  new Set(
    [MAIN, ...secondaryStories(reviewer)].flatMap((story) =>
      blocksOfStory(reviewer, story).map((block) => block.id),
    ),
  );

const drawOperations = (
  flow: Flow,
  blocks: readonly Block[],
  pick: Picker | ((type: string) => Picker) | undefined,
): Operation[] => {
  if (flow.kind === "collisions" && flow.random.chance(0.5)) {
    const names = Object.keys(COLLISIONS);
    const collision = COLLISIONS[flow.random.pick(names)]?.(blocks);
    if (collision && swarmIncludesBatch(flow.swarm, collision)) return collision;
  }
  const count = 1 + flow.random.int(3);
  const operations: Operation[] = [];
  for (let index = 0; index < count; index += 1) {
    const operation = randomOperation(
      flow.generation === "legacy" ? blocksOf(flow) : blocks,
      flow.mode,
      flow.random,
      {
        pick,
        operationHits: flow.operationHits,
        ...(flow.swarm === undefined ? {} : { types: flow.swarm }),
      },
    );
    if (operation) operations.push(operation);
  }
  return operations;
};

const tool = (flow: Flow, name: string, args: unknown) =>
  executeFolioToolCallUntyped(
    name,
    args,
    createReviewerBridge(flow.reviewer, { mode: flow.mode }),
    {},
  );

type Receipt = { applied: readonly { id: string }[] };

/** The operations a receipt says applied, in request order. */
const appliedOf = (operations: readonly Operation[], receipt: Receipt | null) => {
  const ids = new Set(receipt?.applied.map(({ id }) => id));
  return {
    applied: operations.filter((operation) => ids.has(String(operation["id"]))),
    attempted: operations,
  };
};

/**
 * `suggest_changes` with ids on its operations, checked against what they
 * asked; returns how many applied.
 */
const suggestChecked = async (
  flow: Flow,
  args: { operations: unknown },
  entry: string,
): Promise<number> => {
  const operations = Array.isArray(args.operations)
    ? (args.operations as unknown[]).map((operation, index) =>
        typeof operation === "object" && operation !== null && !("id" in operation)
          ? Object.assign({ id: `s-${index + 1}` }, operation)
          : operation,
      )
    : args.operations;
  const done = record(flow, `${entry} ${JSON.stringify({ ...args, operations })}`);
  const pre = await capture(flow.reviewer, flow.mode, { step: flow.session });
  const result = tool(flow, "suggest_changes", { ...args, operations });
  const receipt = result.ok ? (result.result as Receipt) : null;
  done(receipt ? `applied ${receipt.applied.length}` : `refused: ${result.ok ? "" : result.error}`);
  const asked = Array.isArray(operations) ? (operations as Operation[]) : [];
  touch(flow, asked);
  const outcome = appliedOf(asked, receipt);
  await assertRequestedOutcome(flow.reviewer, pre, outcome, `step ${flow.log.length - 1}`);
  return outcome.applied.length;
};

/** Record what a step is about to do; `outcome` completes the entry. */
const record = (flow: Flow, entry: string) => {
  flow.log.push(entry);
  return (outcome: string): void => {
    flow.log[flow.log.length - 1] = `${entry} → ${outcome}`;
  };
};

/** The stories besides the body. */
const secondaryStories = (reviewer: Reviewer): FolioDocumentStoryHandle[] =>
  reviewer
    .listStories()
    .map((story) => story.handle)
    .filter((handle) => handle.type !== "main");

/** One core batch, checked against what it asked; returns how many operations applied. */
const coreBatchStep = async (
  flow: Flow,
  story: FolioDocumentStoryHandle,
  operations: readonly Operation[],
  entry: string,
): Promise<number> => {
  const batch = { ...coreBatch(operations, flow.mode), atomic: flow.random.chance(0.5) };
  const done = record(flow, `${entry} (atomic: ${batch.atomic}) ${JSON.stringify(operations)}`);
  const pre = await capture(flow.reviewer, flow.mode, { story, step: flow.session });
  const result =
    story.type === "main"
      ? flow.reviewer.applyDocumentOperations(batch as never)
      : flow.reviewer.applyDocumentOperationsToStory({ story, batch: batch as never });
  done(`applied ${result.applied.length}, skipped ${result.skipped.length}`);
  touch(flow, operations);
  const outcome = appliedOf(batch.operations, result);
  await assertRequestedOutcome(flow.reviewer, pre, outcome, `step ${flow.log.length - 1}`);
  return outcome.applied.length;
};

/**
 * What a step came to: a batch and how many of its operations applied, a
 * model mistake (refused by design), a step with nothing to act on, or a
 * review or session step.
 */
export type StepEffect =
  | { type: "batch"; applied: number }
  | { type: "mistake" }
  | { type: "skipped" }
  | { type: "ran" };

const SKIPPED = { type: "skipped" } as const satisfies StepEffect;
const RAN = { type: "ran" } as const satisfies StepEffect;

/**
 * One step: drawn from the flow's generator, or the one a flow file planned,
 * whose generator resumes at the position it recorded.
 */
const step = async (flow: Flow, planned?: FlowStep): Promise<StepEffect> => {
  for (const block of blocksOf(flow)) flow.seenIds.add(block.id);
  let action: Action;
  if (planned === undefined) {
    action = flow.random.pick(flow.generation === "legacy" ? LEGACY_ACTIONS : TARGETED_ACTIONS);
  } else {
    action = planned.action;
    flow.random = createRandom(planned.seed);
  }
  flow.planned = planned;
  flow.trace.push({ action, seed: flow.random.state() });
  const { random } = flow;
  switch (action) {
    case "suggest_changes": {
      const applied = await suggestChecked(flow, { operations: randomOperations(flow) }, action);
      return { type: "batch", applied };
    }
    case "core batch": {
      const applied = await coreBatchStep(flow, MAIN, randomOperations(flow), action);
      return { type: "batch", applied };
    }
    case "story batch": {
      const stories = secondaryStories(flow.reviewer);
      if (stories.length === 0) {
        const applied = await coreBatchStep(flow, MAIN, randomOperations(flow), "core batch");
        return { type: "batch", applied };
      }
      const story = random.pick(stories);
      const blocks = blocksOfStory(flow.reviewer, story) as Block[];
      const operations = randomOperations(flow, blocks, pickerFor(flow, story));
      const applied = await coreBatchStep(
        flow,
        story,
        operations,
        `story batch in ${JSON.stringify(story)}`,
      );
      return { type: "batch", applied };
    }
    case "mistake": {
      const [name, build] = random.pick(Object.entries(MISTAKES));
      const args = build(blocksOf(flow), random, [...flow.seenIds]);
      if (typeof args === "object" && args !== null && "operations" in args) {
        await suggestChecked(flow, args, `${action} ${name}`);
        return { type: "mistake" };
      }
      const done = record(flow, `${action} ${name} ${JSON.stringify(args)}`);
      const result = tool(flow, "suggest_changes", args);
      done(result.ok ? "ok" : "refused");
      return { type: "mistake" };
    }
    case "add_comment": {
      const blocks = blocksOf(flow).filter((block) => block.text.length > 0);
      if (blocks.length === 0) {
        record(flow, `${action} skipped`);
        return SKIPPED;
      }
      const pick = pickerFor(flow);
      const picker = typeof pick === "function" ? pick("commentOnBlock") : pick;
      const block = picker?.block(blocks) ?? random.pick(blocks);
      const done = record(flow, `${action} on ${block.id}`);
      const result = tool(flow, "add_comment", { blockId: block.id, text: sentence(random) });
      done(result.ok ? "ok" : result.error);
      return RAN;
    }
    case "reply and resolve": {
      const comments = flow.reviewer.getComments();
      if (comments.length === 0) {
        record(flow, `${action} skipped`);
        return SKIPPED;
      }
      const comment = random.pick(comments);
      record(flow, `${action} ${comment.id}`);
      tool(flow, "reply_comment", { commentId: String(comment.id), text: sentence(random) });
      tool(flow, "resolve_comment", { commentId: String(comment.id), reopen: comment.done });
      return RAN;
    }
    case "accept one":
    case "reject one": {
      const changes = flow.reviewer.getChanges();
      if (changes.length === 0) {
        record(flow, `${action} skipped`);
        return SKIPPED;
      }
      const change = random.pick(changes);
      const done = record(flow, `${action} ${change.type} ${change.id}`);
      const resolved =
        action === "accept one"
          ? flow.reviewer.acceptChange(change)
          : flow.reviewer.rejectChange(change);
      done(String(resolved));
      return RAN;
    }
    case "accept all":
    case "reject all": {
      const done = record(flow, action);
      const accept = action === "accept all";
      const expected = await captureResolution(
        flow.reviewer,
        flow.mode,
        accept ? "accept" : "reject",
      );
      done(String(accept ? flow.reviewer.acceptAll() : flow.reviewer.rejectAll()));
      await assertResolvedTo(flow.reviewer, expected, `step ${flow.log.length - 1}`);
      return RAN;
    }
    case "save and reopen": {
      record(flow, action);
      const { reopened } = await saveAndReopen(flow.reviewer, "save and reopen", saveOptions(flow));
      flow.reviewer = reopened;
      flow.session = "reopened";
      return RAN;
    }
    case "new reviewer": {
      // Another person opens the saved file and carries on under their name.
      record(flow, action);
      const { bytes } = await saveAndReopen(flow.reviewer, action, saveOptions(flow));
      flow.reviewer = await openReviewer(bytes, SECOND_REVIEWER);
      flow.session = "newReviewer";
      return RAN;
    }
    case "selective save": {
      // The patching save the editor uses; the same reviewer keeps working after it.
      const done = record(flow, action);
      const result = await flow.reviewer.save({ repack: "refuse" });
      done(result.type === "selective" ? "selective" : `refused: ${result.reason}`);
      if (result.type !== "selective") return RAN;
      const reopened = await openReviewer(new Uint8Array(result.buffer));
      if (saveOptions(flow).compare !== false) {
        assertSameState(reopened, flow.reviewer, `step ${flow.log.length - 1}: selective save`);
      }
      return RAN;
    }
  }
};

export const SECOND_REVIEWER = "Second Reviewer";

const assertSameState = (reopened: Reviewer, reviewer: Reviewer, context: string): void => {
  const got = JSON.stringify(visibleState(reopened));
  const want = JSON.stringify(visibleState(reviewer));
  if (got !== want) {
    throw new Error(
      `${context}: the reopened package shows something else than the reviewer that saved it\n  expected ${want}\n  got      ${got}`,
    );
  }
};

/**
 * `"suggested"` edits stay out of the package until accepted, so while any
 * are pending the saved package is not what the reviewer shows; it must still
 * save, reopen and read alike.
 */
const saveOptions = (flow: Flow) =>
  flow.mode === "suggested" && flow.reviewer.getChanges().length > 0 ? { compare: false } : {};

// ---------------------------------------------------------------------------
// Step checks
// ---------------------------------------------------------------------------

/** What a check after a step sees. `saved` saves and reopens once per step, however many checks ask. */
export type StepContext = {
  flow: Flow;
  /** The step's index in the flow. */
  index: number;
  /** `step <index>`, for messages. */
  label: string;
  saved: () => Promise<{ bytes: Uint8Array; reopened: Reviewer }>;
};

/** A check every flow runs after every step; throw to fail the flow. */
export type StepCheck = { name: string; check: (context: StepContext) => Promise<void> };

/** The texts of every story besides the body, by handle. */
const storyTexts = (reviewer: Reviewer): Record<string, string[]> =>
  Object.fromEntries(
    secondaryStories(reviewer).map((handle) => [
      JSON.stringify(handle),
      blocksOfStory(reviewer, handle).map((block) => block.text),
    ]),
  );

/**
 * The checks after every step, in order. Add one here to run it in every
 * flow; `saved()` shares the step's save.
 */
export const STEP_CHECKS: StepCheck[] = [
  {
    name: "the package saves and reopens to what the reviewer shows",
    check: async ({ saved }) => {
      await saved();
    },
  },
  {
    name: "every reader of the saved package agrees",
    check: async ({ saved, label }) => assertReadersAgree((await saved()).bytes, label),
  },
  {
    // Headers, footers and notes are not in `visibleState`; their blocks
    // reopen as the reviewer shows them. Suggestions stay out of the package.
    name: "every header, footer and note reopens as the reviewer shows it",
    check: async ({ flow, saved, label }) => {
      if (flow.mode === "suggested") return;
      const { reopened } = await saved();
      const want = storyTexts(flow.reviewer);
      const got = storyTexts(reopened);
      if (JSON.stringify(got) !== JSON.stringify(want)) {
        throw new Error(
          `${label}: a story reopens as something else than the reviewer showed\n  expected ${JSON.stringify(want)}\n  got      ${JSON.stringify(got)}`,
        );
      }
    },
  },
];

// ---------------------------------------------------------------------------
// Running a flow
// ---------------------------------------------------------------------------

/** The fixtures a flow of `kind` and `generation` picks from. */
const flowFixtures = (
  kind: FlowKind,
  generation: Generation,
): Record<string, () => Promise<Uint8Array>> => {
  if (generation === "legacy") {
    // Readers write a merged-cell table as HTML, which the reader comparison
    // does not parse; the collision scenarios cover that fixture.
    return kind === "random" ? FIXTURES : { ...FIXTURES, emoji: COLLISION_FIXTURES.emoji };
  }
  return { ...FIXTURES, emoji: COLLISION_FIXTURES.emoji, ...STORY_FIXTURES };
};

export type FlowOptions = {
  generation?: Generation;
  /** Explicit fixtures are never added to the ordinary seeded fixture selection. */
  fixture?: PublicCorpusFixture | typeof LARGE_DOCUMENT_FIXTURE;
  swarm?: "enabled" | "disabled";
};

/** The fixture and mode a seed's flow runs on. */
export const describeFlow = (
  seed: number,
  kind: FlowKind = "random",
  { generation = "targeted" }: FlowOptions = {},
): { fixture: string; mode: Mode } => {
  const random = createRandom(seed);
  const fixture = random.pick(Object.keys(flowFixtures(kind, generation)));
  return { fixture, mode: random.pick(MODES) };
};

/**
 * A finished flow: what it ran, as a flow file, what each step came to, its
 * signature if asked for and its final saved package if asked for.
 */
export type FlowRun = {
  flow: FlowFile;
  effects: StepEffect[];
  signature: string[];
  saved?: Uint8Array;
};

/**
 * The steps of a run that did nothing: a batch none of whose operations
 * applied, or a step that found nothing to act on. A checked-in flow with
 * one passes without exercising what it guards, as one whose ids drifted
 * from the fixture does.
 */
export const vacuousSteps = ({ effects }: FlowRun): number[] =>
  effects.flatMap((effect, index) => {
    switch (effect.type) {
      case "batch":
        return effect.applied === 0 ? [index] : [];
      case "skipped":
        return [index];
      case "mistake":
      case "ran":
        return [];
      default: {
        const unhandled: never = effect;
        throw new Error(`Unhandled step effect ${JSON.stringify(unhandled)}`);
      }
    }
  });

/** A flow that failed; `flow` holds its steps up to the one that failed. */
export class FlowError extends Error {
  flow: FlowFile;
  constructor(message: string, options: { cause: unknown; flow: FlowFile }) {
    super(message, { cause: options.cause });
    this.flow = options.flow;
  }
}

export type RunOptions = {
  /** Collect the flow's signature as it runs: coverage cells, step outcomes, structures. */
  signature?: boolean;
  /** Save the finished flow and return the package, as a sample. */
  captureSaved?: boolean;
};

/** A few buckets, so a count reads as a shape rather than a number. */
const bucket = (count: number): string => {
  if (count < 2) return String(count);
  return count < 4 ? "2-3" : "4+";
};

/**
 * The document's structure after a step, coarse enough to repeat: which
 * block kinds and target features the body has, which change types are
 * pending and roughly how many comments there are.
 */
const structureOf = (reviewer: Reviewer): string => {
  const features = new Set<string>();
  for (const found of featureIndex(reviewer).features.values()) {
    for (const feature of found) features.add(feature);
  }
  const kinds = new Set((reviewer.getContent() as Block[]).map((block) => String(block.kind)));
  const changes = reviewer.getChanges();
  const changeTypes = new Set(changes.map((change) => String(change.type)));
  return [
    [...kinds].sort().join("+"),
    [...features].sort().join("+"),
    `changes ${[...changeTypes].sort().join("+") || "none"} ${bucket(changes.length)}`,
    `comments ${bucket(reviewer.getComments().length)}`,
  ].join(" | ");
};

/** What a step's log entry says came of it, without its generated values. */
const outcomeOf = (entry: string): string => {
  const arrow = entry.lastIndexOf(" → ");
  return arrow === -1 ? "done" : normalizeAssertion(entry.slice(arrow + 3));
};

type Plan = {
  seed: number;
  kind: FlowKind;
  generation: Generation;
  fixture: string;
  mode: Mode;
  random: Random;
  steps: number;
  planned?: readonly FlowStep[];
  swarm?: string[];
  weights?: FeatureCoverage;
  /** Where the flow came from, for its flow file. */
  origin?: string;
};

const execute = async (plan: Plan, options: RunOptions): Promise<FlowRun> => {
  registerFeatureOperations(Object.keys(GENERATORS));
  const { seed, kind, generation, fixture, mode } = plan;
  if (
    plan.swarm !== undefined &&
    (plan.swarm.length === 0 ||
      new Set(plan.swarm).size !== plan.swarm.length ||
      plan.swarm.some((type) => !Object.hasOwn(GENERATORS, type) || !supports(type, mode)))
  ) {
    throw new TypeError("flow file: swarm contains an unknown or unsupported operation kind");
  }
  const load = (() => {
    if (fixture === LARGE_DOCUMENT_FIXTURE) return largeDocument;
    if (isPublicCorpusFixture(fixture)) return () => loadPublicCorpusFixture(fixture);
    return flowFixtures(kind, generation)[fixture];
  })();
  if (load === undefined) {
    throw new TypeError(`${kind} flow (${generation}): no fixture ${fixture}`);
  }
  const bytes = await load();
  const weights = plan.weights === undefined ? undefined : parseFeatureCoverage(plan.weights);
  const flow: Flow = {
    reviewer: await openReviewer(bytes),
    mode,
    random: plan.random,
    seenIds: new Set(),
    log: [],
    kind,
    generation,
    recent: [],
    session: "fresh",
    trace: [],
    planned: undefined,
    // Read from a reviewer of its own, so the flow's never serves an extra read.
    fixtureIds: blockIdsOf(await openReviewer(bytes)),
    swarm: plan.swarm,
    weights,
    operationHits: weights === undefined ? undefined : featureOperationHits(weights),
  };
  const { log } = flow;
  const file = (): FlowFile => ({
    version: 1,
    kind,
    generation,
    fixture,
    mode,
    seed,
    steps: flow.trace,
    ...(plan.swarm === undefined ? {} : { swarm: plan.swarm }),
    ...(weights === undefined ? {} : { weights }),
    ...(plan.origin === undefined ? {} : { origin: plan.origin }),
  });
  const signature = new Set<string>();
  let captured: Uint8Array | undefined;
  const stopObserving = options.signature
    ? observeHits((key, applied) => signature.add(`cell ${key} ${applied ? "applied" : "refused"}`))
    : () => {};
  // Metamorphic relations (support/metamorphic.ts), checked after every step.
  const relations = await startRelations({ fixture: bytes, reviewer: flow.reviewer, mode, seed });
  const checks: StepCheck[] = [
    ...STEP_CHECKS,
    {
      name: "metamorphic relations",
      check: async ({ flow: current, saved, label }) =>
        relations.afterStep(current.reviewer, await saved(), label),
    },
  ];
  const effects: StepEffect[] = [];
  try {
    for (let index = 0; index < plan.steps; index += 1) {
      effects.push(await step(flow, plan.planned?.[index]));
      const label = `step ${index}`;
      let saved: Promise<{ bytes: Uint8Array; reopened: Reviewer }> | undefined;
      const context: StepContext = {
        flow,
        index,
        label,
        saved: () => (saved ??= saveAndReopen(flow.reviewer, label, saveOptions(flow))),
      };
      for (const { check } of checks) await check(context);
      if (options.signature) {
        const action = flow.trace.at(-1)?.action ?? "";
        signature.add(`step ${action}: ${outcomeOf(log.at(-1) ?? "")}`);
        signature.add(`structure ${structureOf(flow.reviewer)}`);
      }
    }
    await relations.finish();
    if (options.captureSaved) {
      captured = (await saveAndReopen(flow.reviewer, "final sample", saveOptions(flow))).bytes;
    }
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    throw new FlowError(
      [
        `${kind} flow (${generation}) with seed ${seed} (${fixture} / ${mode}) failed at step ${log.length - 1}:`,
        ...log.map((entry, index) => `  ${index}. ${entry}`),
        message,
      ].join("\n"),
      { cause: error, flow: file() },
    );
  } finally {
    stopObserving();
  }
  return {
    flow: file(),
    effects,
    signature: [...signature].sort(),
    ...(captured === undefined ? {} : { saved: captured }),
  };
};

/** Run one seeded flow of `steps` steps; a failure names the step and its log. */
export const runFlow = (
  seed: number,
  steps: number,
  kind: FlowKind = "random",
  {
    generation = "targeted",
    fixture: explicitFixture,
    swarm = process.env["FOLIO_SCENARIO_SWARM"] === "1" ? "enabled" : "disabled",
    ...options
  }: FlowOptions & RunOptions = {},
): Promise<FlowRun> => {
  const random = createRandom(seed);
  const drawnFixture = random.pick(Object.keys(flowFixtures(kind, generation)));
  const fixture = explicitFixture ?? drawnFixture;
  const mode = random.pick(MODES);
  const enabled =
    swarm === "enabled"
      ? drawSwarm(
          seed,
          Object.keys(GENERATORS).filter((type) => supports(type, mode)),
        )
      : undefined;
  const weights = processFeatureWeights();
  return execute(
    {
      seed,
      kind,
      generation,
      fixture,
      mode,
      random,
      steps,
      ...(enabled === undefined ? {} : { swarm: enabled }),
      ...(weights === undefined ? {} : { weights }),
      origin: `${kind} flow seed ${seed}`,
    },
    options,
  );
};

/** Replay a flow file step by step; a failure names the step and its log. */
export const runFlowFile = (file: FlowFile, options: RunOptions = {}): Promise<FlowRun> => {
  const mode = MODES.find((candidate) => candidate === file.mode);
  if (mode === undefined) {
    return Promise.reject(new TypeError(`flow file: unknown mode ${file.mode}`));
  }
  return execute(
    {
      seed: file.seed,
      kind: file.kind,
      generation: file.generation,
      fixture: file.fixture,
      mode,
      random: createRandom(file.seed),
      steps: file.steps.length,
      planned: file.steps,
      ...(file.swarm === undefined ? {} : { swarm: file.swarm }),
      ...(file.weights === undefined ? {} : { weights: file.weights }),
      ...(file.origin === undefined ? {} : { origin: file.origin }),
    },
    options,
  );
};
