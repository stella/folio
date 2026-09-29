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
 */

import { createReviewerBridge, executeFolioToolCallUntyped } from "@stll/folio-agents";
import type { FolioDocumentStoryHandle } from "@stll/folio-core/server";

import type { StepKind } from "./coverage.ts";
import { COLLISION_FIXTURES, FIXTURES, openReviewer, STORY_FIXTURES } from "./documents.ts";
import { assertReadersAgree, saveAndReopen, visibleState } from "./invariants.ts";
import { startRelations } from "./metamorphic.ts";
import { assertRequestedOutcome, assertResolvedTo, capture, captureResolution } from "./oracle.ts";
import {
  type Block,
  COLLISIONS,
  coreBatch,
  MISTAKES,
  MODES,
  type Mode,
  type Operation,
  randomOperation,
} from "./operations.ts";
import { createRandom, type Random, sentence } from "./random.ts";
import { biasedPicker, blocksOfStory, featureIndex, type Picker } from "./targets.ts";

type Reviewer = Awaited<ReturnType<typeof openReviewer>>;

export type FlowKind = "random" | "collisions";

/** How a flow draws its steps; see the module comment. */
export type Generation = "targeted" | "legacy";

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
};

const MAIN: FolioDocumentStoryHandle = { type: "main" };

const blocksOf = (flow: Flow): Block[] => flow.reviewer.getContent() as Block[];

/** The picker a targeted flow aims with, over `story`'s blocks as they are now. */
const pickerFor = (flow: Flow, story: FolioDocumentStoryHandle = MAIN): Picker | undefined =>
  flow.generation === "legacy"
    ? undefined
    : biasedPicker(flow.random, {
        index: featureIndex(flow.reviewer, story),
        recent: flow.recent,
      });

/** Remember the blocks `operations` aimed at, for the next steps to aim near. */
const touch = (flow: Flow, operations: readonly Operation[]): void => {
  for (const operation of operations) {
    const range = operation["range"] as { blockId?: unknown } | undefined;
    const id = operation["blockId"] ?? range?.blockId;
    if (typeof id === "string") flow.recent.push(id);
  }
  flow.recent.splice(0, Math.max(0, flow.recent.length - 12));
};

const randomOperations = (
  flow: Flow,
  blocks: readonly Block[] = blocksOf(flow),
  pick: Picker | undefined = pickerFor(flow),
): Operation[] => {
  if (flow.kind === "collisions" && flow.random.chance(0.5)) {
    const names = Object.keys(COLLISIONS);
    const collision = COLLISIONS[flow.random.pick(names)]?.(blocks);
    if (collision) return collision;
  }
  const count = 1 + flow.random.int(3);
  const operations: Operation[] = [];
  for (let index = 0; index < count; index += 1) {
    const operation = randomOperation(
      flow.generation === "legacy" ? blocksOf(flow) : blocks,
      flow.mode,
      flow.random,
      undefined,
      pick,
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

/** `suggest_changes` with ids on its operations, checked against what they asked. */
const suggestChecked = async (flow: Flow, args: { operations: unknown }, entry: string) => {
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
  await assertRequestedOutcome(
    flow.reviewer,
    pre,
    appliedOf(asked, receipt),
    `step ${flow.log.length - 1}`,
  );
};

/** Record what a step is about to do; `outcome` completes the entry. */
const record = (flow: Flow, entry: string) => {
  flow.log.push(entry);
  return (outcome: string): void => {
    flow.log[flow.log.length - 1] = `${entry} → ${outcome}`;
  };
};

const LEGACY_ACTIONS = [
  "suggest_changes",
  "suggest_changes",
  "core batch",
  "mistake",
  "add_comment",
  "reply and resolve",
  "accept one",
  "reject one",
  "accept all",
  "reject all",
  "save and reopen",
] as const;

/**
 * A targeted flow's steps: the legacy ones, edits aimed at a header, footer
 * or note, and two more session changes (one step in five changes session).
 */
const TARGETED_ACTIONS = [
  ...LEGACY_ACTIONS,
  "story batch",
  "story batch",
  "new reviewer",
  "selective save",
] as const;

type Action = (typeof TARGETED_ACTIONS)[number];

/** The stories besides the body. */
const secondaryStories = (reviewer: Reviewer): FolioDocumentStoryHandle[] =>
  reviewer
    .listStories()
    .map((story) => story.handle)
    .filter((handle) => handle.type !== "main");

/** One core batch, checked against what it asked. */
const coreBatchStep = async (
  flow: Flow,
  story: FolioDocumentStoryHandle,
  operations: readonly Operation[],
  entry: string,
): Promise<void> => {
  const batch = { ...coreBatch(operations, flow.mode), atomic: flow.random.chance(0.5) };
  const done = record(flow, `${entry} (atomic: ${batch.atomic}) ${JSON.stringify(operations)}`);
  const pre = await capture(flow.reviewer, flow.mode, { story, step: flow.session });
  const result =
    story.type === "main"
      ? flow.reviewer.applyDocumentOperations(batch as never)
      : flow.reviewer.applyDocumentOperationsToStory({ story, batch: batch as never });
  done(`applied ${result.applied.length}, skipped ${result.skipped.length}`);
  touch(flow, operations);
  await assertRequestedOutcome(
    flow.reviewer,
    pre,
    appliedOf(batch.operations, result),
    `step ${flow.log.length - 1}`,
  );
};

/** One random step. */
const step = async (flow: Flow): Promise<void> => {
  for (const block of blocksOf(flow)) flow.seenIds.add(block.id);
  const { random } = flow;
  const action: Action = random.pick(
    flow.generation === "legacy" ? LEGACY_ACTIONS : TARGETED_ACTIONS,
  );
  switch (action) {
    case "suggest_changes": {
      await suggestChecked(flow, { operations: randomOperations(flow) }, action);
      return;
    }
    case "core batch": {
      await coreBatchStep(flow, MAIN, randomOperations(flow), action);
      return;
    }
    case "story batch": {
      const stories = secondaryStories(flow.reviewer);
      if (stories.length === 0) {
        await coreBatchStep(flow, MAIN, randomOperations(flow), "core batch");
        return;
      }
      const story = random.pick(stories);
      const blocks = blocksOfStory(flow.reviewer, story) as Block[];
      const operations = randomOperations(flow, blocks, pickerFor(flow, story));
      await coreBatchStep(flow, story, operations, `story batch in ${JSON.stringify(story)}`);
      return;
    }
    case "mistake": {
      const [name, build] = random.pick(Object.entries(MISTAKES));
      const args = build(blocksOf(flow), random, [...flow.seenIds]);
      if (typeof args === "object" && args !== null && "operations" in args) {
        await suggestChecked(flow, args, `${action} ${name}`);
        return;
      }
      const done = record(flow, `${action} ${name} ${JSON.stringify(args)}`);
      const result = tool(flow, "suggest_changes", args);
      done(result.ok ? "ok" : "refused");
      return;
    }
    case "add_comment": {
      const blocks = blocksOf(flow).filter((block) => block.text.length > 0);
      if (blocks.length === 0) {
        record(flow, `${action} skipped`);
        return;
      }
      const block = pickerFor(flow)?.block(blocks) ?? random.pick(blocks);
      const done = record(flow, `${action} on ${block.id}`);
      const result = tool(flow, "add_comment", { blockId: block.id, text: sentence(random) });
      done(result.ok ? "ok" : result.error);
      return;
    }
    case "reply and resolve": {
      const comments = flow.reviewer.getComments();
      if (comments.length === 0) {
        record(flow, `${action} skipped`);
        return;
      }
      const comment = random.pick(comments);
      record(flow, `${action} ${comment.id}`);
      tool(flow, "reply_comment", { commentId: String(comment.id), text: sentence(random) });
      tool(flow, "resolve_comment", { commentId: String(comment.id), reopen: comment.done });
      return;
    }
    case "accept one":
    case "reject one": {
      const changes = flow.reviewer.getChanges();
      if (changes.length === 0) {
        record(flow, `${action} skipped`);
        return;
      }
      const change = random.pick(changes);
      const done = record(flow, `${action} ${change.type} ${change.id}`);
      const resolved =
        action === "accept one"
          ? flow.reviewer.acceptChange(change)
          : flow.reviewer.rejectChange(change);
      done(String(resolved));
      return;
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
      return;
    }
    case "save and reopen": {
      record(flow, action);
      const { reopened } = await saveAndReopen(flow.reviewer, "save and reopen", saveOptions(flow));
      flow.reviewer = reopened;
      flow.session = "reopened";
      return;
    }
    case "new reviewer": {
      // Another person opens the saved file and carries on under their name.
      record(flow, action);
      const { bytes } = await saveAndReopen(flow.reviewer, action, saveOptions(flow));
      flow.reviewer = await openReviewer(bytes, SECOND_REVIEWER);
      flow.session = "newReviewer";
      return;
    }
    case "selective save": {
      // The patching save the editor uses; the same reviewer keeps working after it.
      const done = record(flow, action);
      const result = await flow.reviewer.save({ repack: "refuse" });
      done(result.type === "selective" ? "selective" : `refused: ${result.reason}`);
      if (result.type !== "selective") return;
      const reopened = await openReviewer(new Uint8Array(result.buffer));
      if (saveOptions(flow).compare !== false) {
        assertSameState(reopened, flow.reviewer, `step ${flow.log.length - 1}: selective save`);
      }
      return;
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

export type FlowOptions = { generation?: Generation; captureSaved?: boolean };

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

/** Run one seeded flow of `steps` steps; a failure names the step and its log. */
export const runFlow = async (
  seed: number,
  steps: number,
  kind: FlowKind = "random",
  { generation = "targeted", captureSaved = false }: FlowOptions = {},
): Promise<Uint8Array | undefined> => {
  const random = createRandom(seed);
  const fixtures = flowFixtures(kind, generation);
  const fixture = random.pick(Object.keys(fixtures));
  const mode = random.pick(MODES);
  const bytes = await (fixtures[fixture] as () => Promise<Uint8Array>)();
  const flow: Flow = {
    reviewer: await openReviewer(bytes),
    mode,
    random,
    seenIds: new Set(),
    log: [],
    kind,
    generation,
    recent: [],
    session: "fresh",
  };
  const { log } = flow;
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
  try {
    for (let index = 0; index < steps; index += 1) {
      await step(flow);
      const label = `step ${index}`;
      let saved: Promise<{ bytes: Uint8Array; reopened: Reviewer }> | undefined;
      const context: StepContext = {
        flow,
        index,
        label,
        saved: () => (saved ??= saveAndReopen(flow.reviewer, label, saveOptions(flow))),
      };
      for (const { check } of checks) await check(context);
    }
    await relations.finish();
    if (captureSaved) {
      return (await saveAndReopen(flow.reviewer, "final sample", saveOptions(flow))).bytes;
    }
    return undefined;
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    throw new Error(
      [
        `${kind} flow (${generation}) with seed ${seed} (${fixture} / ${mode}) failed at step ${log.length - 1}:`,
        ...log.map((entry, index) => `  ${index}. ${entry}`),
        message,
      ].join("\n"),
      { cause: error },
    );
  }
};
