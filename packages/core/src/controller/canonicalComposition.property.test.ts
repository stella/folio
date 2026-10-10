import { afterAll, beforeAll, expect, jest, setDefaultTimeout, test } from "bun:test";
import { GlobalRegistrator } from "@happy-dom/global-registrator";
import fc from "fast-check";
import { Schema } from "prosemirror-model";
import { EditorState, TextSelection } from "prosemirror-state";
import type { Transaction } from "prosemirror-state";
import type { Node as ProseMirrorNode } from "prosemirror-model";
import { DOCUMENT_SHAPES, shapeArrayBuffer } from "../__tests__/documentShapes";
import { parseDocx } from "../docx/parser";
import { splitsSurrogatePair } from "../ai-edits/character-boundaries";
import { createCanonicalComposition } from "./canonicalComposition";
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
import { createDocx } from "../docx/rezip";
import { prepareCanonicalDocxInput } from "../docx/canonicalSessionInput";
import { serializeCanonicalSave } from "../docx/canonicalSave";
import { RUN_IDENTITY_MARK_NAME, runIdentityAttrs } from "../prosemirror/runIdentity";

setDefaultTimeout(propertyTestTimeout(60_000));
beforeAll(() => GlobalRegistrator.register());
afterAll(() => GlobalRegistrator.unregister());

const createLateFinalRig = (
  initialText = "alpha",
  source = createEmptyDocument({ initialText }),
) => {
  const paragraph = source.package.document.content.at(0);
  if (!paragraph || paragraph.type !== "paragraph") throw new TypeError("Missing seed paragraph");
  paragraph.paraId ??= "12345678";
  paragraph.textId ??= "87654321";
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
  let nativeMarks = view.state.storedMarks;
  const begin = () => {
    nativeMarks =
      view.state.storedMarks ?? view.state.selection.$from.marksAcross(view.state.selection.$to);
    view.dom.dispatchEvent(new Event("compositionstart", { bubbles: true }));
  };
  return {
    begin,
    boundary,
    session,
    view,
    refusals,
    history: (direction: "undo" | "redo") =>
      view.dom.dispatchEvent(
        new KeyboardEvent("keydown", {
          key: "z",
          ctrlKey: true,
          shiftKey: direction === "redo",
          bubbles: true,
          cancelable: true,
        }),
      ),
    destroy: () => {
      boundary.reset();
      view.destroy();
      mount.remove();
    },
    start: (text: string) => {
      begin();
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
      // Native IME retains its carrier formatting even across an empty update.
      const transaction =
        current === "" ? view.state.tr.setStoredMarks(nativeMarks) : view.state.tr;
      // The observer can emit a minimal diff rather than replace the full range.
      if (text.startsWith(current))
        view.dispatch(transaction.insertText(text.slice(current.length), 1 + current.length));
      else view.dispatch(transaction.insertText(text, 1, 1 + current.length));
      view.dom.dispatchEvent(
        new InputEvent("input", { bubbles: true, inputType: "insertFromComposition", data: text }),
      );
    },
  };
};

type NativeDeletionOptions = {
  view: EditorView;
  inputType: string;
} & ({ type: "withoutTarget" } | { type: "caret" | "selected"; from: number; to: number });

const nativeDeletion = (options: NativeDeletionOptions) => {
  const event = new InputEvent("beforeinput", {
    bubbles: true,
    cancelable: true,
    inputType: options.inputType,
    isComposing: false,
  });
  if (options.type === "withoutTarget") {
    Object.defineProperty(event, "getTargetRanges", { value: () => [] });
    return event;
  }
  const start = options.view.domAtPos(options.from);
  const end = options.view.domAtPos(options.to);
  const target = {
    startContainer: start.node,
    startOffset: start.offset,
    endContainer: end.node,
    endOffset: end.offset,
    collapsed: options.from === options.to,
  } satisfies StaticRange;
  Object.defineProperty(event, "getTargetRanges", { value: () => [target] });
  const selection = document.getSelection();
  if (!selection) throw new TypeError("Missing native selection");
  if (options.type === "caret") {
    const caret = options.view.domAtPos(options.view.state.selection.head);
    selection.collapse(caret.node, caret.offset);
    return event;
  }
  const selected = document.createRange();
  selected.setStart(target.startContainer, target.startOffset);
  selected.setEnd(target.endContainer, target.endOffset);
  selection.removeAllRanges();
  selection.addRange(selected);
  return event;
};

test("native composition cancellation preserves the captured selection and journal", () => {
  jest.useFakeTimers();
  try {
    assertProperty(
      fc.property(
        fc.array(fc.constantFrom("契", "😀", "مرحبا", "alpha", "é"), {
          minLength: 1,
          maxLength: 4,
        }),
        fc.integer({ min: 1, max: 6 }),
        fc.integer({ min: 1, max: 6 }),
        (updates, anchor, head) => {
          for (const inputType of ["deleteContentBackward", "deleteContentForward"]) {
            for (const phase of ["provisional", "refused"]) {
              for (const history of ["empty", "undo", "redo"] as const) {
                const rig = createLateFinalRig();
                try {
                  if (history !== "empty") {
                    const commit = rig.session
                      .prepareReplace(rig.view.state, {
                        from: 1,
                        to: 6,
                        text: "omega",
                        semantic: "typing",
                      })
                      .unwrap();
                    rig.view.updateState(
                      publishCanonicalProjection({
                        session: rig.session,
                        state: rig.view.state,
                        commit,
                      }).unwrap().state,
                    );
                    if (history === "redo") rig.history("undo");
                  }
                  // Refused cancellation targets the restored non-empty selection.
                  let selectedHead = head;
                  if (phase === "refused" && anchor === head)
                    selectedHead = head === 6 ? 5 : head + 1;
                  rig.view.dispatch(
                    rig.view.state.tr.setSelection(
                      TextSelection.create(rig.view.state.doc, anchor, selectedHead),
                    ),
                  );
                  const baseline = rig.view.state;
                  const document = rig.session.document;
                  const version = rig.session.version;
                  const canUndo = rig.session.canUndo;
                  const canRedo = rig.session.canRedo;
                  rig.view.dom.dispatchEvent(new Event("compositionstart", { bubbles: true }));
                  let to = baseline.selection.to;
                  for (const text of updates) {
                    const transaction = rig.view.state.tr
                      .insertText(text, baseline.selection.from, to)
                      .setMeta("composition", 1);
                    if (phase === "refused")
                      transaction.addMark(
                        baseline.selection.from,
                        baseline.selection.from + text.length,
                        canonicalSchema.mark("bold"),
                      );
                    rig.view.dispatch(transaction);
                    to =
                      phase === "refused"
                        ? baseline.selection.to
                        : baseline.selection.from + text.length;
                  }
                  expect(rig.view.composing).toBe(true);
                  // Chromium cancels through a non-composing deletion before
                  // compositionend; it must not delete the restored selection.
                  const cancel = nativeDeletion({
                    view: rig.view,
                    inputType,
                    type: "selected",
                    from: baseline.selection.from,
                    to,
                  });
                  rig.view.dom.dispatchEvent(cancel);
                  jest.advanceTimersByTime(26);
                  expect(cancel.defaultPrevented).toBe(true);
                  expect(rig.view.composing).toBe(false);
                  expect(rig.boundary.isComposing).toBe(false);
                  expect(rig.view.state.doc.eq(baseline.doc)).toBe(true);
                  expect(rig.view.state.selection.eq(baseline.selection)).toBe(true);
                  expect(rig.session.document).toEqual(document);
                  expect(rig.session.version).toBe(version);
                  expect(rig.session.canUndo).toBe(canUndo);
                  expect(rig.session.canRedo).toBe(canRedo);
                  expect(rig.refusals).toHaveLength(phase === "refused" ? 1 : 0);
                  if (history !== "empty") {
                    rig.history(history);
                    expect(rig.view.state.doc.textContent).toBe(
                      history === "undo" ? "alpha" : "omega",
                    );
                    rig.history(history === "undo" ? "redo" : "undo");
                    expect(rig.session.document).toEqual(document);
                  }
                } finally {
                  rig.destroy();
                }
              }
            }
          }
        },
      ),
      {
        numRuns: 20,
        id: "native composition cancellation preserves the captured selection and journal",
      },
    );
  } finally {
    jest.useRealTimers();
  }
});

test("ordinary deletion after native compositionend commits and undoes as a new gesture", () => {
  jest.useFakeTimers();
  try {
    assertProperty(
      fc.property(fc.constantFrom("契約", "alphabet", "😀", "مرحبا", "é"), (text) => {
        for (const delay of [0, 24, 26]) {
          for (const inputType of ["deleteContentBackward", "deleteContentForward"]) {
            const forward = inputType === "deleteContentForward";
            const rig = createLateFinalRig(forward ? "alphax" : "alpha");
            try {
              const original = rig.session.document;
              rig.view.dom.dispatchEvent(new Event("compositionstart", { bubbles: true }));
              rig.view.dispatch(
                rig.view.state.tr
                  .insertText(forward ? text : `${text}x`, 1, 6)
                  .setMeta("composition", 1),
              );
              rig.view.dom.dispatchEvent(new Event("compositionend", { bubbles: true }));
              jest.advanceTimersByTime(delay);
              rig.view.dom.dispatchEvent(
                new InputEvent("beforeinput", { bubbles: true, cancelable: true, inputType }),
              );
              jest.advanceTimersByTime(26);
              expect(rig.view.state.doc.textContent).toBe(text);
              expect(rig.refusals).toEqual([]);
              rig.history("undo");
              expect(rig.view.state.doc.textContent).toBe(`${text}x`);
              rig.history("undo");
              expect(rig.session.document).toEqual(original);
              rig.history("redo");
              expect(rig.view.state.doc.textContent).toBe(`${text}x`);
              rig.history("redo");
              expect(rig.view.state.doc.textContent).toBe(text);
            } finally {
              rig.destroy();
            }
          }
        }
      }),
      {
        numRuns: 10,
        id: "ordinary deletion after native compositionend commits and undoes as a new gesture",
      },
    );
  } finally {
    jest.useRealTimers();
  }
});

test("ordinary deletion recovers a missing native compositionend as a new gesture", () => {
  jest.useFakeTimers();
  try {
    assertProperty(
      fc.property(fc.constantFrom("契約", "alphabet", "😀", "مرحبا", "é", ""), (text) => {
        for (const origin of ["keyboard", "inputOnly"]) {
          for (const target of ["absent", "character", "selectedPartial"]) {
            // A selected whole one-character proposal is the cancellation shape.
            // Ordinary deletion uses a caret or selects only part of a longer proposal.
            if (target === "selectedPartial" && text === "") continue;
            for (const inputType of ["deleteContentBackward", "deleteContentForward"]) {
              const forward = inputType === "deleteContentForward";
              const rig = createLateFinalRig(forward ? "alphax" : "alpha");
              try {
                const original = rig.session.document;
                rig.view.dom.dispatchEvent(new Event("compositionstart", { bubbles: true }));
                rig.view.dispatch(
                  rig.view.state.tr
                    .insertText(forward ? text : `${text}x`, 1, 6)
                    .setMeta("composition", 1),
                );
                const keydown = new KeyboardEvent("keydown", {
                  key: forward ? "Delete" : "Backspace",
                  bubbles: true,
                  cancelable: true,
                  isComposing: false,
                });
                if (origin === "keyboard") {
                  rig.view.dom.dispatchEvent(keydown);
                  expect(keydown.defaultPrevented).toBe(false);
                }
                const deletion = nativeDeletion(
                  target === "absent"
                    ? { view: rig.view, inputType, type: "withoutTarget" }
                    : {
                        view: rig.view,
                        inputType,
                        type: target === "selectedPartial" ? "selected" : "caret",
                        from: 1 + text.length,
                        to: 2 + text.length,
                      },
                );
                rig.view.dom.dispatchEvent(deletion);
                jest.advanceTimersByTime(26);
                expect(
                  rig.view.state.doc.textContent,
                  `${origin}/${target}/${inputType}/${JSON.stringify(text)}`,
                ).toBe(text);
                expect(rig.refusals).toEqual([]);
                rig.history("undo");
                expect(
                  rig.view.state.doc.textContent,
                  `${origin}/${target}/${inputType}/${JSON.stringify(text)}`,
                ).toBe(`${text}x`);
                rig.history("undo");
                expect(rig.session.document).toEqual(original);
                rig.history("redo");
                expect(
                  rig.view.state.doc.textContent,
                  `${origin}/${target}/${inputType}/${JSON.stringify(text)}`,
                ).toBe(`${text}x`);
                rig.history("redo");
                expect(
                  rig.view.state.doc.textContent,
                  `${origin}/${target}/${inputType}/${JSON.stringify(text)}`,
                ).toBe(text);
              } finally {
                rig.destroy();
              }
            }
          }
        }
      }),
      {
        numRuns: 10,
        id: "ordinary deletion recovers a missing native compositionend as a new gesture",
      },
    );
  } finally {
    jest.useRealTimers();
  }
});

test("word, line and cut deletion remain explicit refusals after missing-end recovery", () => {
  jest.useFakeTimers();
  try {
    assertProperty(
      fc.property(fc.constantFrom("契約", "😀", "مرحبا", "é"), (text) => {
        for (const inputType of [
          "deleteWordBackward",
          "deleteWordForward",
          "deleteSoftLineBackward",
          "deleteSoftLineForward",
          "deleteByCut",
        ]) {
          const rig = createLateFinalRig();
          try {
            const original = rig.session.document;
            rig.view.dom.dispatchEvent(new Event("compositionstart", { bubbles: true }));
            rig.view.dispatch(rig.view.state.tr.insertText(text, 1, 6).setMeta("composition", 1));
            const deletion = nativeDeletion({
              view: rig.view,
              inputType,
              type: "selected",
              from: 1,
              to: 1 + text.length,
            });
            rig.view.dom.dispatchEvent(deletion);
            jest.advanceTimersByTime(26);
            expect(deletion.defaultPrevented).toBe(true);
            expect(rig.view.composing).toBe(false);
            expect(rig.boundary.isComposing).toBe(false);
            expect(rig.view.state.doc.textContent).toBe(text);
            expect(rig.refusals.map(({ message }) => message)).toEqual([
              `Input ${inputType} is unavailable in this session.`,
            ]);
            rig.history("undo");
            expect(rig.session.document).toEqual(original);
            expect(rig.session.canUndo).toBe(false);
            rig.history("redo");
            expect(rig.view.state.doc.textContent).toBe(text);
          } finally {
            rig.destroy();
          }
        }
      }),
      {
        numRuns: 10,
        id: "word, line and cut deletion remain explicit refusals after missing-end recovery",
      },
    );
  } finally {
    jest.useRealTimers();
  }
});

test("in-composition deletion edits the native proposal without cancelling its baseline", () => {
  jest.useFakeTimers();
  try {
    assertProperty(
      fc.property(fc.constantFrom("契約", "alphabet", "😀", "مرحبا", "é"), (text) => {
        for (const inputType of ["deleteContentBackward", "deleteContentForward"]) {
          const rig = createLateFinalRig();
          try {
            const original = rig.session.document;
            rig.view.dom.dispatchEvent(new Event("compositionstart", { bubbles: true }));
            const transaction = rig.view.state.tr.insertText(`${text}x`, 1, 6);
            if (inputType === "deleteContentForward")
              transaction.setSelection(TextSelection.create(transaction.doc, 1 + text.length));
            rig.view.dispatch(transaction.setMeta("composition", 1));
            const deletion = new InputEvent("beforeinput", {
              bubbles: true,
              cancelable: true,
              inputType,
              isComposing: true,
            });
            rig.view.dom.dispatchEvent(deletion);
            expect(deletion.defaultPrevented).toBe(false);
            expect(rig.boundary.isComposing).toBe(true);
            rig.view.dispatch(rig.view.state.tr.delete(1 + text.length, 2 + text.length));
            expect(rig.view.state.doc.textContent).toBe(text);
            expect(rig.boundary.isComposing).toBe(true);
            rig.view.dom.dispatchEvent(new Event("compositionend", { bubbles: true }));
            jest.advanceTimersByTime(26);
            expect(rig.session.document.package.document.content).not.toEqual(
              original.package.document.content,
            );
            expect(rig.view.state.doc.textContent).toBe(text);
            expect(rig.refusals).toEqual([]);
            rig.history("undo");
            expect(rig.session.document).toEqual(original);
            rig.history("redo");
            expect(rig.view.state.doc.textContent).toBe(text);
          } finally {
            rig.destroy();
          }
        }
      }),
      {
        numRuns: 10,
        id: "in-composition deletion edits the native proposal without cancelling its baseline",
      },
    );
  } finally {
    jest.useRealTimers();
  }
});

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
      {
        numRuns: 15,
        id: "fake-clock final input before and after expiry is idempotent and undoes as one gesture",
      },
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
        // The same refused gesture reports once; the handled return above
        // proves stale native authorization was consumed.
        expect(rig.refusals).toHaveLength(1);
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
    {
      numRuns: 20,
      id: "every canonical composition exit ends native composition and admits DOM redo",
    },
  );
});

// Absent from every document shape, so a positive-control edit never shares text with its range.
const EDIT_TEXT = "☃";

type TextRange = { from: number; to: number };

const plainTextRanges = (doc: ProseMirrorNode): TextRange[] => {
  const ranges: TextRange[] = [];
  doc.descendants((node, pos) => {
    if (node.text === undefined) return;
    ranges.push({ from: pos, to: pos + node.text.length });
  });
  return ranges;
};

const loadProjections = async () => {
  const projections = [];
  for (const shape of DOCUMENT_SHAPES) {
    const source = await parseDocx(await shapeArrayBuffer(shape), {
      preloadFonts: false,
      detectVariables: false,
    });
    const session = createCanonicalSession(source);
    if (session.isOk()) projections.push({ id: shape.id, doc: session.value.projection.doc });
  }
  return projections;
};

const compose = (baseline: EditorState, nativeEdit: (state: EditorState) => Transaction) => {
  const inputs: { from: number; to: number; text: string }[] = [];
  const refusals: string[] = [];
  const mount = document.createElement("div");
  document.body.append(mount);
  const view = new EditorView(mount, { state: baseline });
  const composition = createCanonicalComposition({
    begin: () => true,
    end: () => {},
    replace: ({ from, to, text }) => inputs.push({ from, to, text }),
    refuse: (reason) => refusals.push(reason),
  });
  composition.start(view);
  composition.accept(view, nativeEdit(view.state).setMeta("composition", 1));
  composition.recover(view);
  const state = view.state;
  view.destroy();
  mount.remove();
  return { inputs, refusals, state };
};

const loadCarrierSource = async () => {
  const source = createEmptyDocument({ initialText: "alpha😀café東京" });
  const paragraph = source.package.document.content.at(0);
  if (paragraph?.type !== "paragraph") throw new TypeError("Missing carrier paragraph");
  const run = paragraph.content.at(0);
  if (run?.type !== "run") throw new TypeError("Missing carrier run");
  run.preservedAttributes = [{ name: "rsidR", value: "00A1B2C3" }];
  const outside = createEmptyDocument({ initialText: "outside" }).package.document.content.at(0);
  if (outside?.type !== "paragraph") throw new TypeError("Missing outside paragraph");
  const outsideRun = outside.content.at(0);
  if (outsideRun?.type !== "run") throw new TypeError("Missing outside run");
  outsideRun.preservedAttributes = [{ name: "rsidR", value: "00D4E5F6" }];
  source.package.document.content.push(outside);
  const input = (await prepareCanonicalDocxInput(await createDocx(source))).unwrap();
  return parseDocx(input, { preloadFonts: false });
};

const omitCarrierIdentity = (state: EditorState, text: string) => {
  const paragraph = state.doc.firstChild;
  if (paragraph === null) throw new TypeError("Missing carrier paragraph");
  const marks = state.selection.$from
    .marks()
    .filter((mark) => mark.type.name !== RUN_IDENTITY_MARK_NAME);
  // Saved native transaction shape: the second update replaces the entire
  // carrier, retaining semantic formatting but omitting its source identity.
  return state.tr.replaceWith(
    1,
    paragraph.content.size + 1,
    state.schema.text(`${text}😀café東京`, marks),
  );
};

test("native carrier identity omission preserves the canonical suffix and history", async () => {
  const source = await loadCarrierSource();
  jest.useFakeTimers();
  try {
    assertProperty(
      fc.property(
        fc.constantFrom("café 東京 é", "shall", "契約", "😀", "alpha"),
        fc.constantFrom("commit", "cancel"),
        (text, completion) => {
          const rig = createLateFinalRig("alpha😀café東京", source);
          const baseline = rig.view.state;
          const original = structuredClone(rig.session.document);
          const identity = baseline.doc
            .nodeAt(1)
            ?.marks.find((mark) => mark.type.name === RUN_IDENTITY_MARK_NAME);
          if (identity === undefined) throw new TypeError("Missing source identity");
          expect(identity?.attrs.preservedAttributes).toEqual([
            { name: "rsidR", value: "00A1B2C3" },
          ]);
          try {
            rig.view.dom.dispatchEvent(new Event("compositionstart", { bubbles: true }));
            rig.view.dispatch(
              rig.view.state.tr.insertText("shall", 1, 6).setMeta("composition", 1),
            );
            rig.view.dispatch(omitCarrierIdentity(rig.view.state, text).setMeta("composition", 1));
            expect(rig.session.projection.doc.eq(baseline.doc)).toBe(true);
            expect(
              rig.view.state.doc
                .nodeAt(1 + text.length)
                ?.marks.some((mark) => mark.type.name === RUN_IDENTITY_MARK_NAME),
            ).toBe(false);
            if (completion === "cancel") rig.boundary.reset();
            else {
              rig.view.dom.dispatchEvent(new Event("compositionend", { bubbles: true }));
              jest.advanceTimersByTime(26);
            }
            expect(rig.refusals).toEqual([]);
            const changed = completion === "commit" && text !== "alpha";
            expect(rig.session.canUndo).toBe(changed);
            if (!changed) {
              expect(structuredClone(rig.session.document)).toEqual(original);
              expect(rig.view.state.doc.eq(baseline.doc)).toBe(true);
              return;
            }
            expect(rig.view.state.doc.firstChild?.textContent).toBe(`${text}😀café東京`);
            const direct = createCanonicalSession(source).unwrap();
            const directState = EditorState.create({
              doc: direct.projection.doc,
              selection: TextSelection.create(direct.projection.doc, 1, 6),
            });
            const expected = publishCanonicalProjection({
              session: direct,
              state: directState,
              commit: direct
                .prepareReplace(directState, { from: 1, to: 6, text, semantic: "composition" })
                .unwrap(),
            }).unwrap().state;
            expect(rig.view.state.doc.eq(expected.doc)).toBe(true);
            expect(structuredClone(rig.session.document)).toEqual(structuredClone(direct.document));
            expect(
              rig.view.state.doc
                .nodeAt(1 + text.length)
                ?.marks.find((mark) => mark.type.name === RUN_IDENTITY_MARK_NAME)?.attrs
                .preservedAttributes,
            ).toEqual(identity.attrs.preservedAttributes);
            const committed = structuredClone(rig.session.document);
            const suffix = committed.package.document.content.at(0);
            expect(
              suffix?.type === "paragraph" &&
                suffix.content.some(
                  (run) =>
                    run.type === "run" &&
                    run.preservedAttributes?.some((attr) => attr.value === "00A1B2C3"),
                ),
            ).toBe(true);
            rig.history("undo");
            expect(structuredClone(rig.session.document)).toEqual(original);
            expect(rig.view.state.doc.eq(baseline.doc)).toBe(true);
            rig.history("redo");
            expect(structuredClone(rig.session.document)).toEqual(committed);
          } finally {
            rig.destroy();
          }
        },
      ),
      {
        examples: [
          ["café 東京 é", "cancel"],
          ["café 東京 é", "commit"],
          ["alpha", "commit"],
        ],
        id: "native carrier identity omission preserves the canonical suffix and history",
      },
    );
  } finally {
    jest.useRealTimers();
  }
});

test("native carrier saves equal direct canonical replacements", async () => {
  const source = await loadCarrierSource();
  await assertProperty(
    fc.asyncProperty(fc.constantFrom("café 東京 é", "契約", "😀"), async (text) => {
      const rig = createLateFinalRig("alpha😀café東京", source);
      try {
        rig.view.dom.dispatchEvent(new Event("compositionstart", { bubbles: true }));
        rig.view.dispatch(rig.view.state.tr.insertText("shall", 1, 6).setMeta("composition", 1));
        rig.view.dispatch(omitCarrierIdentity(rig.view.state, text).setMeta("composition", 1));
        rig.boundary.handleDOMEvents.blur(rig.view);
        expect(rig.refusals).toEqual([]);
        const direct = createCanonicalSession(source).unwrap();
        const state = EditorState.create({ doc: direct.projection.doc });
        publishCanonicalProjection({
          session: direct,
          state,
          commit: direct
            .prepareReplace(state, { from: 1, to: 6, text, semantic: "composition" })
            .unwrap(),
        }).unwrap();
        const nativeSave = await serializeCanonicalSave({
          snapshot: rig.session.captureSaveSnapshot(),
        });
        const directSave = await serializeCanonicalSave({ snapshot: direct.captureSaveSnapshot() });
        const nativeDoc = await parseDocx(nativeSave.buffer, { preloadFonts: false });
        const directDoc = await parseDocx(directSave.buffer, { preloadFonts: false });
        expect(structuredClone(nativeDoc.package.document)).toEqual(
          structuredClone(directDoc.package.document),
        );
      } finally {
        rig.destroy();
      }
    }),
    { examples: [["café 東京 é"]], id: "native carrier saves equal direct canonical replacements" },
  );
});

test("native carrier identity tolerance refuses every other mark or attribute drift", async () => {
  const session = createCanonicalSession(await loadCarrierSource()).unwrap();
  const baseline = EditorState.create({
    doc: session.projection.doc,
    selection: TextSelection.create(session.projection.doc, 1, 6),
  });
  assertProperty(
    fc.property(
      fc.constantFrom(
        "differentIdentity",
        "differentPayload",
        "outsideIdentity",
        "bold",
        "paragraphAttribute",
      ),
      (drift) => {
        const native = compose(baseline, (state) => {
          const transaction = omitCarrierIdentity(state, "shall");
          const identity = state.schema.marks[RUN_IDENTITY_MARK_NAME];
          if (identity === undefined) throw new TypeError("Missing identity mark");
          switch (drift) {
            case "differentIdentity":
              return transaction.addMark(1, 6, identity.create(runIdentityAttrs(999)));
            case "differentPayload":
              return transaction.addMark(
                1,
                6,
                identity.create(
                  runIdentityAttrs(0, {
                    preservedAttributes: [{ name: "rsidR", value: "FFFFFFFF" }],
                  }),
                ),
              );
            case "outsideIdentity": {
              const outside = transaction.doc.firstChild;
              if (outside === null) throw new TypeError("Missing carrier");
              expect(
                identity.isInSet(transaction.doc.nodeAt(outside.nodeSize + 1)?.marks ?? []),
              ).toBeDefined();
              return transaction.removeMark(
                outside.nodeSize + 1,
                transaction.doc.content.size - 1,
                identity,
              );
            }
            case "bold": {
              const bold = state.schema.marks.bold;
              if (bold === undefined) throw new TypeError("Missing bold mark");
              return transaction.addMark(1, 6, bold.create());
            }
            case "paragraphAttribute":
              return transaction.setNodeMarkup(0, undefined, {
                ...transaction.doc.firstChild?.attrs,
                alignment: "right",
              });
            default: {
              const unexpected: never = drift;
              throw new TypeError(`Unknown native drift: ${unexpected}`);
            }
          }
        });
        expect(native.inputs).toEqual([]);
        expect(native.refusals).toHaveLength(1);
        expect(native.state.doc.eq(baseline.doc)).toBe(true);
        expect(native.state.selection.eq(baseline.selection)).toBe(true);
      },
    ),
    { id: "native carrier identity tolerance refuses every other mark or attribute drift" },
  );
});

test("native carrier commits require rebuilding the captured baseline", async () => {
  const source = await loadCarrierSource();
  jest.useFakeTimers();
  const rig = createLateFinalRig("alpha😀café東京", source);
  const baseline = rig.view.state;
  const updateState = rig.view.updateState.bind(rig.view);
  try {
    rig.view.dom.dispatchEvent(new Event("compositionstart", { bubbles: true }));
    rig.view.dispatch(omitCarrierIdentity(rig.view.state, "café 東京 é").setMeta("composition", 1));
    // Ablate only the baseline rebuild: the canonical session must reject the
    // identity-free provisional carrier as a stale projection, never adopt it.
    rig.view.updateState = (state) => {
      if (state !== baseline) updateState(state);
    };
    expect(() => rig.boundary.handleDOMEvents.blur(rig.view)).toThrow("stale canonical projection");
    expect(rig.session.canUndo).toBe(false);
  } finally {
    rig.view.updateState = updateState;
    rig.destroy();
    jest.useRealTimers();
  }
});

test("a native rewrite of identical text never commits a replacement", async () => {
  const docs = await loadProjections();
  const executed = { markLoss: 0, textEdit: 0 };
  assertProperty(
    fc.property(
      fc.constantFrom(...docs),
      fc.nat(),
      fc.nat(),
      fc.nat(),
      ({ doc }, rangeIndex, startOffset, length) => {
        const ranges = plainTextRanges(doc);
        const range = ranges.at(rangeIndex % ranges.length);
        if (range === undefined) return;
        const from = range.from + (startOffset % (range.to - range.from));
        const to = from + 1 + (length % (range.to - from));
        const rangeText = doc.textBetween(range.from, range.to, "", "");
        if (
          splitsSurrogatePair(rangeText, from - range.from) ||
          splitsSurrogatePair(rangeText, to - range.from)
        )
          return;
        const baseline = EditorState.create({
          doc,
          selection: TextSelection.create(doc, from, to),
        });
        const sameText = doc.textBetween(from, to, "", "");
        const rewrite = (state: EditorState) => state.tr.insertText(sameText, from, to);
        if (!baseline.apply(rewrite(baseline)).doc.eq(doc)) executed.markLoss++;

        const native = compose(baseline, rewrite);
        expect(native.inputs).toEqual([]);
        expect(native.state.doc.eq(doc)).toBe(true);
        expect(native.state.selection.eq(baseline.selection)).toBe(true);

        // Positive control: a real text change through the same path still commits exactly.
        const edit = compose(baseline, (state) => state.tr.insertText(EDIT_TEXT, from, to));
        expect(edit.refusals).toEqual([]);
        expect(edit.inputs).toEqual([{ from, to, text: EDIT_TEXT }]);
        executed.textEdit++;
      },
    ),
    { id: "a native rewrite of identical text never commits a replacement" },
  );
  expect(docs.length).toBeGreaterThan(0);
  expect(executed.markLoss).toBeGreaterThan(0);
  expect(executed.textEdit).toBeGreaterThan(0);
});

test("native composition lifecycle keeps cancelled selections and commits explicit finals", () => {
  jest.useFakeTimers();
  try {
    assertProperty(
      fc.property(
        fc.array(fc.constantFrom("", "shall", "café 東京 é", "😀", "契約"), {
          minLength: 1,
          maxLength: 5,
        }),
        fc.constantFrom("cancel" as const, "commit" as const),
        fc.constantFrom("", "shall", "契約"),
        (updates, completion, finalText) => {
          const rig = createLateFinalRig("alpha😀café東京");
          try {
            const baseline = rig.view.state;
            const original = structuredClone(rig.session.document);
            const marks = baseline.selection.$from.marksAcross(baseline.selection.$to);
            rig.begin();
            let current = "alpha";
            const update = (text: string) => {
              rig.view.dom.dispatchEvent(
                new InputEvent("beforeinput", {
                  bubbles: true,
                  inputType: "insertCompositionText",
                  data: text,
                  isComposing: true,
                }),
              );
              const transaction =
                current === "" ? rig.view.state.tr.setStoredMarks(marks) : rig.view.state.tr;
              rig.view.dispatch(
                transaction.insertText(text, 1, 1 + current.length).setMeta("composition", 1),
              );
              current = text;
              expect(rig.session.canUndo).toBe(false);
            };
            for (const text of updates) update(text);
            if (completion === "cancel") update("");
            else rig.final(finalText, current);
            // The DOM fixture aliases CompositionEvent to Event and drops its data init.
            const end = new CompositionEvent("compositionend", { bubbles: true });
            Object.defineProperty(end, "data", { value: completion === "cancel" ? "" : finalText });
            rig.view.dom.dispatchEvent(end);
            jest.advanceTimersByTime(26);
            expect(rig.refusals).toEqual([]);
            if (completion === "cancel") {
              expect(structuredClone(rig.session.document)).toEqual(original);
              expect(rig.view.state.doc.eq(baseline.doc)).toBe(true);
              expect(rig.view.state.selection.eq(baseline.selection)).toBe(true);
              expect(rig.session.canUndo).toBe(false);
            } else {
              expect(rig.view.state.doc.textContent).toBe(`${finalText}😀café東京`);
              expect(rig.session.canUndo).toBe(true);
              rig.history("undo");
              expect(structuredClone(rig.session.document)).toEqual(original);
              rig.history("redo");
              expect(rig.view.state.doc.textContent).toBe(`${finalText}😀café東京`);
            }
          } finally {
            rig.destroy();
          }
        },
      ),
      {
        examples: [
          [["shall", "café 東京 é", ""], "cancel", ""],
          [["shall", ""], "commit", "契約"],
          [["shall", ""], "commit", ""],
        ],
        id: "native composition lifecycle keeps cancelled selections and commits explicit finals",
      },
    );
  } finally {
    jest.useRealTimers();
  }
});

test("native cancellation restoration requires its end evidence", () => {
  jest.useFakeTimers();
  try {
    for (const evidence of ["nativeEnd", "ablatedEnd"] as const) {
      const rig = createLateFinalRig("alpha😀café東京");
      try {
        const baseline = rig.view.state;
        rig.begin();
        rig.view.dispatch(rig.view.state.tr.insertText("", 1, 6).setMeta("composition", 1));
        const end = new CompositionEvent("compositionend", { bubbles: true });
        Object.defineProperty(end, "data", { value: "" });
        if (evidence === "ablatedEnd") {
          const handleEnd = rig.boundary.handleDOMEvents.compositionend;
          rig.boundary.handleDOMEvents.compositionend = (view) =>
            handleEnd(view, new Event("compositionend"));
        }
        rig.view.dom.dispatchEvent(end);
        jest.advanceTimersByTime(26);
        expect(rig.refusals).toEqual([]);
        expect(rig.view.state.doc.eq(baseline.doc)).toBe(evidence === "nativeEnd");
        expect(rig.view.state.doc.textContent).toBe(
          evidence === "nativeEnd" ? "alpha😀café東京" : "😀café東京",
        );
        expect(rig.session.canUndo).toBe(evidence === "ablatedEnd");
      } finally {
        rig.destroy();
      }
    }
  } finally {
    jest.useRealTimers();
  }
});

test("classified native finals commit empty replacements during end recovery", () => {
  jest.useFakeTimers();
  try {
    assertProperty(
      fc.property(fc.constantFrom("insertText", "insertReplacementText"), (inputType) => {
        const rig = createLateFinalRig("alpha😀café東京");
        try {
          rig.begin();
          rig.view.dispatch(rig.view.state.tr.insertText("", 1, 6).setMeta("composition", 1));
          const end = new CompositionEvent("compositionend", { bubbles: true });
          Object.defineProperty(end, "data", { value: "" });
          // The boundary receives end evidence before the native owner has flushed it.
          rig.boundary.handleDOMEvents.compositionend(rig.view, end);
          expect(rig.view.composing).toBe(true);
          const final = new InputEvent("beforeinput", {
            bubbles: true,
            cancelable: true,
            inputType,
            data: "",
          });
          rig.view.dom.dispatchEvent(final);
          expect(final.defaultPrevented).toBe(true);
          jest.advanceTimersByTime(26);
          expect(rig.refusals).toEqual([]);
          expect(rig.view.state.doc.textContent).toBe("😀café東京");
          expect(rig.session.canUndo).toBe(true);
          rig.history("undo");
          expect(rig.view.state.doc.textContent).toBe("alpha😀café東京");
        } finally {
          rig.destroy();
        }
      }),
      { id: "classified native finals commit empty replacements during end recovery" },
    );
  } finally {
    jest.useRealTimers();
  }
});
