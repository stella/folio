/**
 * The plugins that carry positions through a transaction do not map every
 * paragraph, or every step's range, through every step.
 *
 * A paste, a replace-all, an accept-all or a document compare dispatches one
 * transaction with a step per change. Walking each paragraph through each of
 * them costs O(paragraphs x steps), which on a 2,200-paragraph compare was
 * most of the time spent applying it. The guard is a count rather than a
 * stopwatch: every position a step map is asked about is one call on
 * `StepMap`, so a per-step walk shows up as paragraphs x steps calls however
 * fast the machine is.
 */

import { afterEach, describe, expect, spyOn, test } from "bun:test";
import { EditorState, type Plugin } from "prosemirror-state";
import { StepMap } from "prosemirror-transform";

import { randomSchema } from "../__tests__/randomTransactions";
import { AutoBidiDetectionExtension } from "./features/AutoBidiDetectionExtension";
import { ParagraphChangeTrackerExtension } from "./features/ParagraphChangeTrackerExtension";
import { ParaIdAllocatorExtension } from "./features/ParaIdAllocatorExtension";
import { RunIdentityExtension } from "./marks/RunIdentityExtension";

const PARAGRAPHS = 600;
const EDITED = 200;

const pluginsOf = (runtime: { plugins?: Plugin[] }): Plugin[] => runtime.plugins ?? [];

const PLUGINS = {
  "paragraph change tracker": pluginsOf(
    ParagraphChangeTrackerExtension().onSchemaReady({ schema: randomSchema }),
  ),
  "paragraph-id allocator": pluginsOf(
    ParaIdAllocatorExtension().onSchemaReady({ schema: randomSchema }),
  ),
  "base-direction detection": pluginsOf(
    AutoBidiDetectionExtension().onSchemaReady({ schema: randomSchema }),
  ),
  "run identity": pluginsOf(RunIdentityExtension().onSchemaReady({ schema: randomSchema })),
};

const largeDocument = () =>
  randomSchema.node(
    "doc",
    null,
    Array.from({ length: PARAGRAPHS }, (_, index) =>
      randomSchema.node("paragraph", { paraId: (index + 1).toString(16).padStart(8, "0") }, [
        randomSchema.text(`Paragraph ${String(index)} of a long contract.`),
      ]),
    ),
  );

/**
 * Words inserted into every third paragraph, a split, a mark and an attribute
 * change: the step shapes a compare batch or a replace-all dispatches.
 */
const manyStepTransaction = (state: EditorState) => {
  const tr = state.tr;
  const starts: number[] = [];
  state.doc.forEach((_node, offset) => starts.push(offset));
  for (let edit = EDITED - 1; edit >= 0; edit--) {
    // SAFETY: 3 * edit < PARAGRAPHS.
    const start = starts[3 * edit]!;
    tr.insertText("amended ", start + 1);
  }
  tr.split(3);
  const bold = randomSchema.marks["bold"];
  if (bold) {
    tr.addMark(10, 60, bold.create());
  }
  tr.setNodeAttribute(0, "direction", { source: "manual" });
  return tr;
};

/** Every question asked of any step map while `plugins` handle one transaction. */
const stepMapCalls = (plugins: readonly Plugin[]): { calls: number; steps: number } => {
  const state = EditorState.create({ doc: largeDocument(), plugins: [...plugins] });
  const tr = manyStepTransaction(state);
  const map = spyOn(StepMap.prototype, "map");
  const mapResult = spyOn(StepMap.prototype, "mapResult");
  state.applyTransaction(tr);
  return { calls: map.mock.calls.length + mapResult.mock.calls.length, steps: tr.steps.length };
};

afterEach(() => {
  // `spyOn` wraps the shared prototype; restore it for the next case.
  // oxlint-disable-next-line typescript/unbound-method -- restoring the spied prototype methods
  (StepMap.prototype.map as unknown as { mockRestore?: () => void }).mockRestore?.();
  // oxlint-disable-next-line typescript/unbound-method -- restoring the spied prototype methods
  (StepMap.prototype.mapResult as unknown as { mockRestore?: () => void }).mockRestore?.();
});

describe("a many-step transaction", () => {
  // ProseMirror itself maps the selection through every step once.
  const baseline = (): number => stepMapCalls([]).calls;

  for (const [name, plugins] of Object.entries(PLUGINS)) {
    test(`costs the ${name} no step-map call per paragraph per step`, () => {
      const without = baseline();
      const { calls, steps } = stepMapCalls(plugins);
      const perStepWalk = PARAGRAPHS * steps;
      // Linear in paragraphs plus steps: each position is asked about only by
      // the steps whose changed range covers it.
      expect(calls - without).toBeLessThan(2 * (PARAGRAPHS + steps));
      expect(calls - without).toBeLessThan(perStepWalk / 50);
    });
  }
});
