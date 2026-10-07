import { afterAll, beforeAll, expect, jest, setDefaultTimeout, test } from "bun:test";
import { GlobalRegistrator } from "@happy-dom/global-registrator";
import fc from "fast-check";
import { Schema } from "prosemirror-model";
import { EditorState, TextSelection } from "prosemirror-state";
import { EditorView } from "prosemirror-view";
import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";
import { CANONICAL_COMPOSITION_INPUT_TYPES, createCanonicalInputBoundary } from "./canonicalInput";
import {
  createCanonicalSession,
  CanonicalSessionError,
  publishCanonicalProjection,
} from "./canonicalSession";
import { CANONICAL_GAP } from "../types/canonicalCapabilities";
import { schema as canonicalSchema } from "../prosemirror/schema";
import { createEmptyDocument } from "../utils/createDocument";

setDefaultTimeout(propertyTestTimeout(30_000));
beforeAll(() => GlobalRegistrator.register());
afterAll(() => GlobalRegistrator.unregister());

const createLateFinalRig = () => {
  const source = createEmptyDocument({ initialText: "alpha" });
  const paragraph = source.package.document.content.at(0);
  if (!paragraph || paragraph.type !== "paragraph") throw new TypeError("Missing seed paragraph");
  paragraph.paraId = "12345678";
  paragraph.textId = "87654321";
  const session = createCanonicalSession(source).unwrap();
  const refusals: CanonicalSessionError[] = [];
  const boundary = createCanonicalInputBoundary({
    beginComposition: () => session.beginComposition().isOk(),
    endComposition: () => session.endComposition(),
    breakUndoGroup: () => session.breakUndoGroup(),
    replace: (input) => {
      const commit = session.prepareReplace(view.state, input).unwrap();
      view.updateState(
        publishCanonicalProjection({ session, state: view.state, commit }).unwrap().state,
      );
    },
    refuse: (message) =>
      refusals.push(
        new CanonicalSessionError({ gap: CANONICAL_GAP.dispatch, message, reason: "refused" }),
      ),
    undo: () => {
      const commit = session.prepareUndo(view.state).unwrap();
      view.updateState(
        publishCanonicalProjection({ session, state: view.state, commit }).unwrap().state,
      );
      return true;
    },
    redo: () => {
      const commit = session.prepareRedo(view.state).unwrap();
      view.updateState(
        publishCanonicalProjection({ session, state: view.state, commit }).unwrap().state,
      );
      return true;
    },
  });
  const doc = session.projection.doc;
  const mount = document.body.appendChild(document.createElement("div"));
  const view = new EditorView(mount, {
    state: EditorState.create({
      schema: canonicalSchema,
      doc,
      selection: TextSelection.create(doc, 1, 6),
    }),
    handleKeyDown: boundary.handleKeyDown,
    handleTextInput: boundary.handleTextInput,
    handleDOMEvents: boundary.handleDOMEvents,
    dispatchTransaction: (transaction) => {
      if (boundary.acceptComposition(view, transaction)) return;
      if (transaction.docChanged) {
        if (!boundary.commitNativeProposal(view, transaction)) boundary.refuseNativeMutation(view);
        return;
      }
      view.updateState(view.state.apply(transaction));
    },
  });
  return {
    boundary,
    session,
    view,
    refusals,
    destroy: () => {
      boundary.reset();
      view.destroy();
      mount.remove();
    },
    start: (text: string) => {
      view.dom.dispatchEvent(new Event("compositionstart", { bubbles: true }));
      view.dispatch(view.state.tr.insertText(text, 1, 6).setMeta("composition", 1));
      view.dom.dispatchEvent(new Event("compositionend", { bubbles: true }));
    },
    final: (text: string, current: string) => {
      const event = new InputEvent("beforeinput", {
        bubbles: true,
        cancelable: true,
        inputType: "insertFromComposition",
        data: text,
      });
      view.dom.dispatchEvent(event);
      expect(event.defaultPrevented).toBe(false);
      // The observer can emit a minimal diff rather than replace the full range.
      if (text.startsWith(current))
        view.dispatch(view.state.tr.insertText(text.slice(current.length), 1 + current.length));
      else view.dispatch(view.state.tr.insertText(text, 1, 1 + current.length));
      view.dom.dispatchEvent(
        new InputEvent("input", { bubbles: true, inputType: "insertFromComposition", data: text }),
      );
    },
  };
};

test("fake-clock final input before and after expiry is idempotent and undoes as one gesture", () => {
  jest.useFakeTimers();
  try {
    assertProperty(
      fc.property(
        fc.constantFrom("契", "😀", "مرحبا", "a", "alpha", ""),
        fc.integer({ min: 1, max: 4 }),
        (provisional, duplicates) => {
          for (const delay of [0, 24, 25, 26, 250, 2500]) {
            for (const final of [provisional, `${provisional}約`]) {
              const rig = createLateFinalRig();
              try {
                const original = rig.session.document;
                const originalSelection = rig.view.state.selection.toJSON();
                rig.start(provisional);
                jest.advanceTimersByTime(delay);
                for (let repeat = 0; repeat < duplicates; repeat++) {
                  const version = rig.session.version;
                  rig.final(final, repeat === 0 ? provisional : final);
                  jest.advanceTimersByTime(26);
                  if (repeat > 0 || (delay >= 25 && final === provisional))
                    expect(rig.session.version).toBe(version);
                }
                expect(rig.refusals).toEqual([]);
                expect(rig.view.state.doc.textContent).toBe(final);
                expect(rig.boundary.isComposing).toBe(false);
                expect(rig.view.composing).toBe(false);
                expect(rig.session.projection.doc.eq(rig.view.state.doc)).toBe(true);
                const composed = rig.session.document;
                const finalSelection = rig.view.state.selection.toJSON();
                if (final !== "alpha") {
                  rig.view.dom.dispatchEvent(
                    new KeyboardEvent("keydown", {
                      key: "z",
                      ctrlKey: true,
                      bubbles: true,
                      cancelable: true,
                    }),
                  );
                  expect(rig.session.document).toEqual(original);
                  expect(rig.view.state.selection.toJSON()).toEqual(originalSelection);
                  expect(rig.session.canUndo).toBe(false);
                  rig.view.dom.dispatchEvent(
                    new KeyboardEvent("keydown", {
                      key: "z",
                      ctrlKey: true,
                      shiftKey: true,
                      bubbles: true,
                      cancelable: true,
                    }),
                  );
                  expect(rig.session.document).toEqual(composed);
                  expect(rig.view.state.selection.toJSON()).toEqual(finalSelection);
                  expect(rig.session.canRedo).toBe(false);
                } else expect(rig.session.canUndo).toBe(false);
              } finally {
                rig.destroy();
              }
            }
          }
        },
      ),
      { numRuns: 15 },
    );
  } finally {
    jest.useRealTimers();
  }
});

test("fake-clock completion receipts admit only final input from the unchanged owner", () => {
  const latePolicy = {
    insertCompositionText: "refuse",
    insertFromComposition: "native",
    deleteCompositionText: "refuse",
    deleteByComposition: "refuse",
  } as const satisfies Record<
    (typeof CANONICAL_COMPOSITION_INPUT_TYPES)[number],
    "native" | "refuse"
  >;
  jest.useFakeTimers();
  try {
    for (const delay of [24, 25, 250]) {
      for (const inputType of CANONICAL_COMPOSITION_INPUT_TYPES) {
        const rig = createLateFinalRig();
        try {
          rig.start("契");
          jest.advanceTimersByTime(delay);
          const event = new InputEvent("beforeinput", { inputType, data: "契", cancelable: true });
          const native = delay < 25 || latePolicy[inputType] === "native";
          expect(rig.boundary.handleDOMEvents.beforeinput(rig.view, event)).toBe(!native);
          expect(event.defaultPrevented).toBe(!native);
          expect(rig.refusals.length).toBe(native ? 0 : 1);
          if (!native) expect(rig.refusals.at(0)?.reason).toBe("refused");
        } finally {
          rig.destroy();
        }
      }
    }
    for (const invalidate of [
      "reset",
      "nextKey",
      "nextInput",
      "selection",
      "otherView",
      "refused",
    ] as const) {
      const rig = createLateFinalRig();
      const other = createLateFinalRig();
      try {
        rig.start("契");
        if (invalidate === "refused") {
          rig.view.dispatch(
            rig.view.state.tr.addMark(1, 2, canonicalSchema.mark("bold")).setMeta("composition", 1),
          );
          expect(rig.refusals).toHaveLength(1);
        }
        jest.advanceTimersByTime(26);
        switch (invalidate) {
          case "reset":
            rig.boundary.reset();
            break;
          case "nextKey":
            rig.boundary.handleKeyDown(
              rig.view,
              new KeyboardEvent("keydown", { key: "ArrowRight" }),
            );
            break;
          case "nextInput":
            rig.boundary.handleDOMEvents.beforeinput(
              rig.view,
              new InputEvent("beforeinput", { inputType: "historyUndo", cancelable: true }),
            );
            break;
          case "selection":
            rig.view.updateState(
              rig.view.state.apply(
                rig.view.state.tr.setSelection(TextSelection.create(rig.view.state.doc, 1)),
              ),
            );
            break;
          case "otherView":
          case "refused":
            break;
          default: {
            const unexpected: never = invalidate;
            throw new TypeError(`Unknown invalidation: ${unexpected}`);
          }
        }
        const version = rig.session.version;
        const document = rig.session.document;
        const event = new InputEvent("beforeinput", {
          inputType: "insertFromComposition",
          data: "契約",
          cancelable: true,
        });
        expect(
          rig.boundary.handleDOMEvents.beforeinput(
            invalidate === "otherView" ? other.view : rig.view,
            event,
          ),
        ).toBe(true);
        expect(event.defaultPrevented).toBe(true);
        expect(rig.refusals.at(-1)?.reason).toBe("refused");
        expect(rig.session.version).toBe(version);
        expect(rig.session.document).toBe(document);
      } finally {
        rig.destroy();
        other.destroy();
      }
    }
  } finally {
    jest.useRealTimers();
  }
});

test("fake-clock state changes after final authorization consume the stale receipt", () => {
  jest.useFakeTimers();
  try {
    for (const delay of [26, 250, 2500]) {
      const rig = createLateFinalRig();
      try {
        rig.start("契");
        jest.advanceTimersByTime(delay);
        const event = new InputEvent("beforeinput", {
          inputType: "insertFromComposition",
          data: "契約",
          cancelable: true,
        });
        expect(rig.boundary.handleDOMEvents.beforeinput(rig.view, event)).toBe(false);
        rig.view.updateState(
          rig.view.state.apply(
            rig.view.state.tr.setSelection(TextSelection.create(rig.view.state.doc, 1)),
          ),
        );
        const version = rig.session.version;
        const document = rig.session.document;
        rig.view.dispatch(rig.view.state.tr.insertText("約", 2));
        expect(rig.refusals).toHaveLength(1);
        expect(rig.refusals.at(-1)?.reason).toBe("refused");
        expect(rig.view.state.doc.textContent).toBe("契");
        // Without consuming stale authorization, text input remains native forever.
        expect(rig.boundary.handleTextInput(rig.view, 1, 1, "x")).toBe(true);
        expect(rig.refusals).toHaveLength(2);
        expect(rig.session.version).toBe(version);
        expect(rig.session.document).toBe(document);
      } finally {
        rig.destroy();
      }
    }
  } finally {
    jest.useRealTimers();
  }
});

test("fake-clock late final cannot authorize marks or another replacement range", () => {
  jest.useFakeTimers();
  try {
    for (const mutation of ["mark", "foreign", "unclassified"] as const) {
      const rig = createLateFinalRig();
      try {
        rig.start("契");
        jest.advanceTimersByTime(26);
        const state = rig.view.state;
        const version = rig.session.version;
        if (mutation !== "unclassified") {
          const event = new InputEvent("beforeinput", {
            inputType: "insertFromComposition",
            data: "契約",
            cancelable: true,
          });
          expect(rig.boundary.handleDOMEvents.beforeinput(rig.view, event)).toBe(false);
        }
        const transaction = state.tr.insertText(mutation === "foreign" ? "other" : "契約", 1, 2);
        if (mutation === "mark") transaction.addMark(1, 3, canonicalSchema.mark("bold"));
        rig.view.dispatch(transaction);
        expect(rig.refusals.at(-1)?.reason).toBe("refused");
        expect(rig.session.version).toBe(version);
        expect(rig.view.state.doc.eq(state.doc)).toBe(true);
      } finally {
        rig.destroy();
      }
    }
  } finally {
    jest.useRealTimers();
  }
});

test("fake-clock corrections keep separate native gestures in separate undo groups", () => {
  jest.useFakeTimers();
  const rig = createLateFinalRig();
  try {
    rig.start("契");
    jest.advanceTimersByTime(26);
    rig.final("契約", "契");
    jest.advanceTimersByTime(26);
    const first = rig.session.document;
    rig.view.dom.dispatchEvent(new Event("compositionstart", { bubbles: true }));
    const caret = rig.view.state.selection.head;
    rig.view.dispatch(rig.view.state.tr.insertText("X", caret).setMeta("composition", 2));
    rig.view.dom.dispatchEvent(new Event("compositionend", { bubbles: true }));
    jest.advanceTimersByTime(26);
    const commit = rig.session.prepareUndo(rig.view.state).unwrap();
    rig.view.updateState(
      publishCanonicalProjection({ session: rig.session, state: rig.view.state, commit }).unwrap()
        .state,
    );
    expect(rig.session.document).toEqual(first);
    expect(rig.view.state.doc.textContent).toBe("契約");
    expect(rig.session.canUndo).toBe(true);
  } finally {
    rig.destroy();
    jest.useRealTimers();
  }
});

test("fake-clock classified late finals survive input before the delayed observer flush", async () => {
  jest.useFakeTimers();
  try {
    for (const delay of [26, 250, 2500]) {
      for (const correction of ["契", "契約"]) {
        const rig = createLateFinalRig();
        try {
          rig.start("契");
          jest.advanceTimersByTime(26);
          const version = rig.session.version;
          const event = new InputEvent("beforeinput", {
            inputType: "insertFromComposition",
            data: correction,
            cancelable: true,
          });
          expect(rig.boundary.handleDOMEvents.beforeinput(rig.view, event)).toBe(false);
          rig.boundary.handleDOMEvents.input(rig.view);
          await Promise.resolve();
          jest.advanceTimersByTime(delay);
          rig.view.dispatch(rig.view.state.tr.insertText(correction, 1, 2));
          expect(rig.refusals).toEqual([]);
          expect(rig.view.state.doc.textContent).toBe(correction);
          expect(rig.session.version).toBe(version + (correction === "契" ? 0 : 1));
          const commit = rig.session.prepareUndo(rig.view.state).unwrap();
          rig.view.updateState(
            publishCanonicalProjection({
              session: rig.session,
              state: rig.view.state,
              commit,
            }).unwrap().state,
          );
          expect(rig.view.state.doc.textContent).toBe("alpha");
          expect(rig.session.canUndo).toBe(false);
        } finally {
          rig.destroy();
        }
      }
    }
  } finally {
    jest.useRealTimers();
  }
});

const schema = new Schema({
  nodes: {
    doc: { content: "paragraph+" },
    paragraph: { content: "text*", toDOM: () => ["p", 0] },
    text: {},
  },
  marks: { strong: { toDOM: () => ["strong", 0] } },
});

// Direct key handler calls bypass PM's composing gate; use actual DOM delivery.
test("every canonical composition exit ends native composition and admits DOM redo", async () => {
  await assertProperty(
    fc.asyncProperty(
      fc.array(fc.constantFrom("契", "😀", "مرحبا", "café"), { minLength: 1, maxLength: 4 }),
      fc.boolean(),
      async (parts, refused) => {
        for (const exit of [
          "commit",
          "refusedFinal",
          "recover",
          "cancel",
          "reset",
          "blur",
          "mousedown",
        ] as const) {
          if (exit === "refusedFinal" && !refused) continue;
          const text = parts.join("");
          let replacements = 0;
          let ends = 0;
          let redos = 0;
          let refusals = 0;
          const boundary = createCanonicalInputBoundary({
            beginComposition: () => true,
            endComposition: () => {
              ends++;
            },
            replace: () => {
              replacements++;
            },
            refuse: () => {
              refusals++;
            },
            undo: () => false,
            redo: () => {
              redos++;
              return true;
            },
          });
          const doc = schema.node("doc", null, [
            schema.node("paragraph", null, schema.text("alpha")),
          ]);
          const baseline = EditorState.create({ doc, selection: TextSelection.create(doc, 1, 6) });
          const mount = document.body.appendChild(document.createElement("div"));
          const view = new EditorView(mount, {
            state: baseline,
            handleKeyDown: boundary.handleKeyDown,
            handleDOMEvents: boundary.handleDOMEvents,
            dispatchTransaction: (transaction) => {
              if (!boundary.acceptComposition(view, transaction))
                view.updateState(view.state.apply(transaction));
            },
          });
          try {
            view.dom.dispatchEvent(new Event("compositionstart", { bubbles: true }));
            expect(view.composing).toBe(true);
            const transaction = view.state.tr.insertText(text, 1, 6).setMeta("composition", 1);
            if (refused) transaction.addMark(1, 1 + text.length, schema.mark("strong"));
            view.dispatch(transaction);
            expect(boundary.isComposing).toBe(true);
            expect(view.composing).toBe(true);
            for (const init of [{ isComposing: true }, { keyCode: 229 }]) {
              for (const key of ["Escape", "z"]) {
                const event = new KeyboardEvent("keydown", {
                  ...init,
                  key,
                  ctrlKey: true,
                  shiftKey: true,
                  bubbles: true,
                  cancelable: true,
                });
                view.dom.dispatchEvent(event);
                expect(event.defaultPrevented).toBe(false);
                expect(boundary.handleKeyDown(view, event)).toBe(false);
                expect(boundary.isComposing).toBe(true);
              }
            }
            switch (exit) {
              case "commit":
                view.dom.dispatchEvent(new Event("compositionend", { bubbles: true }));
                await new Promise<void>((resolve) => setTimeout(resolve, 40));
                break;
              case "refusedFinal": {
                // Chromium's refused commit need not deliver compositionend.
                const final = new InputEvent("beforeinput", {
                  inputType: "insertText",
                  data: text,
                  bubbles: true,
                  cancelable: true,
                });
                view.dom.dispatchEvent(final);
                expect(final.defaultPrevented).toBe(true);
                break;
              }
              case "recover":
                boundary.handleDOMEvents.beforeinput(
                  view,
                  new InputEvent("beforeinput", { inputType: "historyUndo", cancelable: true }),
                );
                break;
              case "cancel":
                view.dom.dispatchEvent(
                  new KeyboardEvent("keydown", { key: "Escape", bubbles: true, cancelable: true }),
                );
                break;
              case "reset":
                boundary.reset();
                break;
              case "blur":
              case "mousedown":
                boundary.handleDOMEvents[exit](view);
                break;
              default: {
                const unexpected: never = exit;
                throw new TypeError(`Unknown exit: ${unexpected}`);
              }
            }
            expect(boundary.isComposing).toBe(false);
            expect(view.composing).toBe(false);
            expect(view.state.doc.eq(baseline.doc)).toBe(true);
            expect(view.state.selection.eq(baseline.selection)).toBe(true);
            expect(ends).toBe(1);
            expect(refusals).toBe(refused ? 1 : 0);
            expect(replacements).toBe(!refused && exit !== "cancel" && exit !== "reset" ? 1 : 0);
            // Native IME keys also bypass extension shortcuts after recovery.
            for (const init of [{ isComposing: true }, { keyCode: 229 }]) {
              const nativeKey = new KeyboardEvent("keydown", {
                ...init,
                key: "z",
                ctrlKey: true,
                shiftKey: true,
                bubbles: true,
                cancelable: true,
              });
              view.dom.dispatchEvent(nativeKey);
              expect(nativeKey.defaultPrevented).toBe(false);
              expect(redos).toBe(0);
            }
            const redo = new KeyboardEvent("keydown", {
              key: "z",
              ctrlKey: true,
              shiftKey: true,
              bubbles: true,
              cancelable: true,
            });
            view.dom.dispatchEvent(redo);
            expect(redo.defaultPrevented).toBe(true);
            expect(redos).toBe(1);
            boundary.reset();
            expect(ends).toBe(1);
          } finally {
            boundary.reset();
            view.destroy();
            mount.remove();
          }
        }
      },
    ),
    { numRuns: 20 },
  );
});
