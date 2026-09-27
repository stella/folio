/**
 * The seeded random flows the fuzz scenarios run: chains of public
 * operations (tool calls, core batches, model mistakes, comments,
 * accept/reject, save and reopen) over the synthetic documents, checking
 * after every step that the document saves, reopens to what the reviewer
 * showed, and reads alike everywhere, and that every applied operation did
 * what it asked (support/oracle.ts). A flow is fully determined by its seed,
 * step count and kind: a `"collisions"` flow draws half its batches from
 * `COLLISIONS` and may run on the emoji fixture too.
 */

import { createReviewerBridge, executeFolioToolCallUntyped } from "@stll/folio-agents";

import { COLLISION_FIXTURES, FIXTURES, openReviewer } from "./documents.ts";
import { assertReadersAgree, saveAndReopen } from "./invariants.ts";
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

type Reviewer = Awaited<ReturnType<typeof openReviewer>>;

type Flow = {
  reviewer: Reviewer;
  mode: Mode;
  random: Random;
  /** Ids read earlier that may no longer name a block. */
  seenIds: Set<string>;
  /** What each step did, written before it runs so a throw still names it. */
  log: string[];
  kind: FlowKind;
};

export type FlowKind = "random" | "collisions";

const blocksOf = (flow: Flow): Block[] => flow.reviewer.getContent() as Block[];

const randomOperations = (flow: Flow): Operation[] => {
  if (flow.kind === "collisions" && flow.random.chance(0.5)) {
    const names = Object.keys(COLLISIONS);
    const collision = COLLISIONS[flow.random.pick(names)]?.(blocksOf(flow));
    if (collision) return collision;
  }
  const count = 1 + flow.random.int(3);
  const operations: Operation[] = [];
  for (let index = 0; index < count; index += 1) {
    const operation = randomOperation(blocksOf(flow), flow.mode, flow.random);
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
  return { applied: operations.filter((operation) => ids.has(String(operation["id"]))) };
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
  const pre = await capture(flow.reviewer, flow.mode);
  const result = tool(flow, "suggest_changes", { ...args, operations });
  const receipt = result.ok ? (result.result as Receipt) : null;
  done(receipt ? `applied ${receipt.applied.length}` : `refused: ${result.ok ? "" : result.error}`);
  await assertRequestedOutcome(
    flow.reviewer,
    pre,
    appliedOf(Array.isArray(operations) ? (operations as Operation[]) : [], receipt),
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

/** One random step. */
const step = async (flow: Flow): Promise<void> => {
  for (const block of blocksOf(flow)) flow.seenIds.add(block.id);
  const { random } = flow;
  const action = random.pick([
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
  ] as const);
  switch (action) {
    case "suggest_changes": {
      await suggestChecked(flow, { operations: randomOperations(flow) }, action);
      return;
    }
    case "core batch": {
      const operations = randomOperations(flow);
      const batch = { ...coreBatch(operations, flow.mode), atomic: random.chance(0.5) };
      const done = record(
        flow,
        `${action} (atomic: ${batch.atomic}) ${JSON.stringify(operations)}`,
      );
      const pre = await capture(flow.reviewer, flow.mode);
      const result = flow.reviewer.applyDocumentOperations(batch as never);
      done(`applied ${result.applied.length}, skipped ${result.skipped.length}`);
      await assertRequestedOutcome(
        flow.reviewer,
        pre,
        appliedOf(batch.operations, result),
        `step ${flow.log.length - 1}`,
      );
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
      const block = random.pick(blocks);
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
      return;
    }
  }
};

/**
 * `"suggested"` edits stay out of the package until accepted, so while any
 * are pending the saved package is not what the reviewer shows; it must still
 * save, reopen and read alike.
 */
const saveOptions = (flow: Flow) =>
  flow.mode === "suggested" && flow.reviewer.getChanges().length > 0 ? { compare: false } : {};

/** The fixtures a flow of `kind` picks from. */
const FLOW_FIXTURES = {
  random: FIXTURES,
  // Readers write a merged-cell table as HTML, which the reader comparison
  // does not parse; the collision scenarios cover that fixture.
  collisions: { ...FIXTURES, emoji: COLLISION_FIXTURES.emoji },
} as const;

/** The fixture and mode a seed's flow runs on. */
export const describeFlow = (
  seed: number,
  kind: FlowKind = "random",
): { fixture: string; mode: Mode } => {
  const random = createRandom(seed);
  const fixture = random.pick(Object.keys(FLOW_FIXTURES[kind]));
  return { fixture, mode: random.pick(MODES) };
};

/** Run one seeded flow of `steps` steps; a failure names the step and its log. */
export const runFlow = async (
  seed: number,
  steps: number,
  kind: FlowKind = "random",
): Promise<void> => {
  const random = createRandom(seed);
  const fixtures: Record<string, () => Promise<Uint8Array>> = FLOW_FIXTURES[kind];
  const fixture = random.pick(Object.keys(fixtures));
  const mode = random.pick(MODES);
  const build = fixtures[fixture] as () => Promise<Uint8Array>;
  const flow: Flow = {
    reviewer: await openReviewer(await build()),
    mode,
    random,
    seenIds: new Set(),
    log: [],
    kind,
  };
  const { log } = flow;
  try {
    for (let index = 0; index < steps; index += 1) {
      await step(flow);
      const { bytes } = await saveAndReopen(flow.reviewer, `step ${index}`, saveOptions(flow));
      await assertReadersAgree(bytes, `step ${index}`);
    }
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    throw new Error(
      [
        `${kind} flow with seed ${seed} (${fixture} / ${mode}) failed at step ${log.length - 1}:`,
        ...log.map((entry, index) => `  ${index}. ${entry}`),
        message,
      ].join("\n"),
      { cause: error },
    );
  }
};
