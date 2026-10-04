/**
 * Replacing a node selection over an inline atom in suggesting mode. Whatever
 * replaces the atom (typed text, a paste, an inserted inline node), accepting
 * every change must read as the edit made directly and rejecting every change
 * must restore the original document.
 */

import { describe, test } from "bun:test";
import fc from "fast-check";
import { Fragment, Slice } from "prosemirror-model";
import type { Node as PMNode } from "prosemirror-model";
import { EditorState, NodeSelection } from "prosemirror-state";
import type { Transaction } from "prosemirror-state";

import { assertProperty, propertyTestTimeout } from "../../../../../test/property-testing";

import { acceptAllChanges, rejectAllChanges } from "../commands/comments";
import { schema } from "../schema";
import { dispatchEditorTextInput } from "../textInput";
import {
  canCarryTrackedRunMark,
  TRACKED_RUN_INLINE_ATOM_DISPOSITIONS,
} from "../trackedRunInlineAtoms";
import { createSuggestionModePlugin } from "./suggestionMode";

/** Every node-selectable inline atom a tracked run can own, from the disposition map. */
const SELECTABLE_ATOMS: readonly PMNode[] = Object.keys(TRACKED_RUN_INLINE_ATOM_DISPOSITIONS)
  .map((name) => schema.nodes[name]?.createAndFill() ?? null)
  .filter(
    (node): node is PMNode =>
      node !== null &&
      node.isInline &&
      node.isAtom &&
      NodeSelection.isSelectable(node) &&
      canCarryTrackedRunMark(node),
  );

type HeadlessView = {
  state: EditorState;
  composing: boolean;
  dispatch: (tr: Transaction) => void;
  someProp: (name: string, call: (handler: never) => unknown) => unknown;
};

const headlessView = (state: EditorState): HeadlessView => {
  const view: HeadlessView = {
    state,
    composing: false,
    dispatch: (tr) => {
      view.state = view.state.apply(tr);
    },
    someProp: (name, call) => {
      for (const plugin of view.state.plugins) {
        const prop: unknown = Reflect.get(plugin.props, name);
        if (typeof prop !== "function") continue;
        // SAFETY: each caller names the prop whose handler signature it passes.
        const result = call(prop.bind(plugin) as never);
        if (result) return result;
      }
      return undefined;
    },
  };
  return view;
};

const paste = (view: HeadlessView, slice: Slice): void => {
  const event = { clipboardData: null, preventDefault: () => undefined };
  const handled = view.someProp("handlePaste", (handler: (...args: unknown[]) => boolean) =>
    handler(view, event, slice),
  );
  if (!handled) {
    view.dispatch(view.state.tr.replaceSelection(slice).setMeta("paste", true));
  }
};

type Replacement = { kind: string; run: (view: HeadlessView) => void };

const letters = fc.string({
  unit: fc.constantFrom("a", "b", "c", " "),
  minLength: 0,
  maxLength: 4,
});

const replacementArb: fc.Arbitrary<Replacement> = fc.oneof(
  fc.string({ unit: fc.constantFrom("x", "y", "z"), minLength: 1, maxLength: 4 }).map((text) => ({
    kind: `type ${JSON.stringify(text)}`,
    run: (view: HeadlessView) => {
      for (const character of text) {
        dispatchEditorTextInput(view, character);
      }
    },
  })),
  fc.constantFrom<Replacement>(
    {
      kind: "paste inline text",
      run: (view) => paste(view, new Slice(Fragment.from(schema.text("Pasted")), 0, 0)),
    },
    {
      kind: "insert an inline node",
      run: (view) => view.dispatch(view.state.tr.replaceSelectionWith(schema.node("hardBreak"))),
    },
  ),
);

const runIn = (doc: PMNode, atomPos: number, suggesting: boolean, replacement: Replacement) => {
  const state = EditorState.create({
    doc,
    plugins: [createSuggestionModePlugin(suggesting, "Author")],
  });
  const view = headlessView(state.apply(state.tr.setSelection(NodeSelection.create(doc, atomPos))));
  replacement.run(view);
  return view.state;
};

const resolveAll = (state: EditorState, command: typeof acceptAllChanges): PMNode => {
  let resolved = state;
  command()(state, (tr) => {
    resolved = state.apply(tr);
  });
  return resolved.doc;
};

describe("suggesting mode replaces a node-selected inline atom as a tracked change", () => {
  test("the atom kinds are derived, not empty", () => {
    if (!SELECTABLE_ATOMS.some((atom) => atom.type.name === "image")) {
      throw new Error("Expected the image atom among the selectable inline atoms");
    }
  });

  test(
    "accept-all reads as the direct edit and reject-all restores the original",
    () => {
      assertProperty(
        fc.property(
          fc.constantFrom(...SELECTABLE_ATOMS),
          letters,
          letters,
          replacementArb,
          (atom, before, after, replacement) => {
            const inline = [
              ...(before ? [schema.text(before)] : []),
              atom,
              ...(after ? [schema.text(after)] : []),
            ];
            const doc = schema.node("doc", null, [schema.node("paragraph", null, inline)]);
            const atomPos = 1 + before.length;
            const direct = runIn(doc, atomPos, false, replacement).doc;
            const suggested = runIn(doc, atomPos, true, replacement);
            const accepted = resolveAll(suggested, acceptAllChanges);
            const rejected = resolveAll(suggested, rejectAllChanges);
            const context = `${atom.type.name}, ${replacement.kind}`;
            if (!accepted.eq(direct)) {
              throw new Error(
                `${context}: accept-all ${JSON.stringify(accepted.toJSON())} ≠ direct ${JSON.stringify(direct.toJSON())}`,
              );
            }
            if (!rejected.eq(doc)) {
              throw new Error(
                `${context}: reject-all ${JSON.stringify(rejected.toJSON())} ≠ original ${JSON.stringify(doc.toJSON())}`,
              );
            }
          },
        ),
        { numRuns: 60 },
      );
    },
    propertyTestTimeout(30_000),
  );
});
