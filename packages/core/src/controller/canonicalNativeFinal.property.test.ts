import { afterAll, beforeAll, expect, jest, test } from "bun:test";
import { GlobalRegistrator } from "@happy-dom/global-registrator";
import fc from "fast-check";
import { EditorState, Plugin, TextSelection } from "prosemirror-state";
import { EditorView } from "prosemirror-view";
import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";
import { schema } from "../prosemirror/schema";
import { createEmptyDocument } from "../utils/createDocument";
import { createCanonicalInputBoundary } from "./canonicalInput";
import { createCanonicalSession, publishCanonicalProjection } from "./canonicalSession";

beforeAll(() => GlobalRegistrator.register());
afterAll(() => GlobalRegistrator.unregister());

const createRig = (refusal: "marks" | "plugin" = "marks") => {
  const source = createEmptyDocument({ initialText: "alpha" });
  const paragraph = source.package.document.content.at(0);
  if (!paragraph || paragraph.type !== "paragraph") throw new TypeError("Missing seed paragraph");
  paragraph.paraId = "12345678";
  paragraph.textId = "87654321";
  const session = createCanonicalSession(source).unwrap();
  const refusals: string[] = [];
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
    refuse: (message) => refusals.push(message),
    undo: () => false,
    redo: () => false,
  });
  const doc = session.projection.doc;
  const mount = document.body.appendChild(document.createElement("div"));
  const view = new EditorView(mount, {
    state: EditorState.create({
      schema,
      doc,
      selection: TextSelection.create(doc, 1, 6),
      plugins:
        refusal === "plugin"
          ? [
              new Plugin({
                filterTransaction: (transaction) => {
                  if (typeof transaction.getMeta("composition") === "number")
                    throw new TypeError("Native composition projection failed");
                  return true;
                },
              }),
            ]
          : [],
    }),
    handleTextInput: boundary.handleTextInput,
    handleDOMEvents: boundary.handleDOMEvents,
    dispatchTransaction: (transaction) => {
      if (!boundary.acceptComposition(view, transaction))
        view.updateState(view.state.apply(transaction));
    },
  });
  return {
    session,
    boundary,
    view,
    refusals,
    destroy: () => {
      boundary.reset();
      view.destroy();
      mount.remove();
    },
  };
};

test(
  "native final text is committed exactly once or refused without a journal change",
  () => {
    jest.useFakeTimers();
    try {
      assertProperty(
        fc.property(
          fc.constantFrom("契", "😀", "مرحبا", "alpha", "alphabet", "alpine"),
          fc.constantFrom("provisional", "refused", "empty"),
          (text, phase) => {
            for (const ended of [false, true]) {
              for (const mutation of ["selection", "diff"] as const) {
                const refusals =
                  phase === "refused" ? (["marks", "plugin"] as const) : (["marks"] as const);
                for (const refusal of refusals) {
                  const rig = createRig(refusal);
                  try {
                    const document = rig.session.document;
                    const baseline = rig.view.state;
                    expect(
                      baseline.doc
                        .nodeAt(1)
                        ?.marks.some((mark) => mark.type.name === "runIdentity"),
                    ).toBe(true);
                    rig.view.dom.dispatchEvent(new Event("compositionstart", { bubbles: true }));
                    if (phase !== "empty") {
                      const provisional =
                        mutation === "diff" && text.startsWith("alpha")
                          ? rig.view.state.tr
                              .insertText(text.slice(5), 6, 6)
                              .setMeta("composition", 1)
                          : rig.view.state.tr.insertText(text, 1, 6).setMeta("composition", 1);
                      if (phase === "refused" && refusal === "marks")
                        provisional.addMark(1, 1 + text.length, schema.mark("bold"));
                      if (phase === "provisional" && mutation === "diff" && text === "alphabet") {
                        expect(
                          provisional.doc
                            .nodeAt(1)
                            ?.marks.some((mark) => mark.type.name === "runIdentity"),
                        ).toBe(true);
                        expect(baseline.tr.insertText(text, 1, 6).doc.eq(provisional.doc)).toBe(
                          false,
                        );
                      }
                      rig.view.dispatch(provisional);
                    }
                    expect(rig.view.composing).toBe(true);
                    if (ended) {
                      rig.view.dom.dispatchEvent(new Event("compositionend", { bubbles: true }));
                      expect(rig.view.composing).toBe(false);
                      expect(rig.boundary.isComposing).toBe(true);
                    }
                    const final = new InputEvent("beforeinput", {
                      bubbles: true,
                      cancelable: true,
                      inputType: ended ? "insertFromComposition" : "insertText",
                      data: text,
                    });
                    // Exercise both native end orderings before the delayed finish;
                    // no cleanup cancel precedes the final input.
                    rig.view.dom.dispatchEvent(final);
                    expect(final.defaultPrevented).toBe(!ended);
                    if (ended) {
                      const to = phase === "provisional" ? 1 + text.length : 6;
                      rig.view.dispatch(
                        rig.view.state.tr.insertText(text, 1, to).setMeta("composition", 2),
                      );
                      jest.advanceTimersByTime(26);
                    }
                    expect(rig.view.composing).toBe(false);
                    expect(rig.boundary.isComposing).toBe(false);
                    expect(rig.refusals).toHaveLength(phase === "refused" ? 1 : 0);
                    expect(rig.view.state.doc.textContent).toBe(
                      phase === "refused" ? "alpha" : text,
                    );
                    const edited = phase !== "refused" && text !== "alpha";
                    expect(rig.session.version).toBe(edited ? 1 : 0);
                    expect(rig.session.projection.doc.eq(rig.view.state.doc)).toBe(true);
                    if (edited) {
                      const undo = rig.session.prepareUndo(rig.view.state).unwrap();
                      rig.view.updateState(
                        publishCanonicalProjection({
                          session: rig.session,
                          state: rig.view.state,
                          commit: undo,
                        }).unwrap().state,
                      );
                    }
                    expect(rig.session.document).toEqual(document);
                    expect(rig.view.state.doc.eq(baseline.doc)).toBe(true);
                    expect(rig.session.canUndo).toBe(false);
                  } finally {
                    rig.destroy();
                  }
                }
              }
            }
          },
        ),
        {
          numRuns: 36,
          id: "native final text is committed exactly once or refused without a journal change",
        },
      );
    } finally {
      jest.useRealTimers();
    }
  },
  propertyTestTimeout(5_000),
);

test.each(["provisional-active", "provisional-ended", "refused-ended"] as const)(
  "a different typing payload after %s recovery remains a separate edit",
  (phase) => {
    const rig = createRig();
    try {
      rig.view.dom.dispatchEvent(new Event("compositionstart", { bubbles: true }));
      const provisional = rig.view.state.tr.insertText("契", 1, 6).setMeta("composition", 1);
      if (phase === "refused-ended") provisional.addMark(1, 2, schema.mark("bold"));
      rig.view.dispatch(provisional);
      if (phase !== "provisional-active")
        rig.view.dom.dispatchEvent(new Event("compositionend", { bubbles: true }));
      rig.view.dom.dispatchEvent(
        new InputEvent("beforeinput", {
          bubbles: true,
          cancelable: true,
          inputType: "insertText",
          data: "NEXT",
        }),
      );
      const refused = phase === "refused-ended";
      expect(rig.refusals).toHaveLength(refused ? 1 : 0);
      expect(rig.view.state.doc.textContent).toBe(refused ? "NEXT" : "契NEXT");
      expect(rig.session.version).toBe(refused ? 1 : 2);
      const undo = rig.session.prepareUndo(rig.view.state).unwrap();
      rig.view.updateState(
        publishCanonicalProjection({
          session: rig.session,
          state: rig.view.state,
          commit: undo,
        }).unwrap().state,
      );
      expect(rig.view.state.doc.textContent).toBe(refused ? "alpha" : "契");
      expect(rig.session.canUndo).toBe(!refused);
    } finally {
      rig.destroy();
    }
  },
);

test(
  "same-payload typing after native end is always a separate gesture",
  () => {
    jest.useFakeTimers();
    try {
      assertProperty(
        fc.property(fc.constantFrom("契", "😀", "مرحبا", "alphabet"), (text) => {
          for (const delay of [0, 24, 25, 26]) {
            for (const ending of ["nativeEnd", "consumedFinal"] as const) {
              const rig = createRig();
              try {
                const original = rig.session.document;
                rig.view.dom.dispatchEvent(new Event("compositionstart", { bubbles: true }));
                rig.view.dispatch(
                  rig.view.state.tr.insertText(text, 1, 6).setMeta("composition", 1),
                );
                if (ending === "nativeEnd") {
                  rig.view.dom.dispatchEvent(new Event("compositionend", { bubbles: true }));
                } else {
                  const final = new InputEvent("beforeinput", {
                    bubbles: true,
                    cancelable: true,
                    inputType: "insertText",
                    data: text,
                  });
                  rig.view.dom.dispatchEvent(final);
                  expect(final.defaultPrevented).toBe(true);
                  expect(rig.session.version).toBe(1);
                }
                expect(rig.view.composing).toBe(false);
                jest.advanceTimersByTime(delay);
                rig.view.dom.dispatchEvent(
                  new InputEvent("beforeinput", {
                    bubbles: true,
                    cancelable: true,
                    inputType: "insertText",
                    data: text,
                  }),
                );
                expect(rig.refusals).toEqual([]);
                expect(rig.view.state.doc.textContent).toBe(text + text);
                expect(rig.session.version).toBe(2);
                expect(rig.session.projection.doc.eq(rig.view.state.doc)).toBe(true);
                const typingUndo = rig.session.prepareUndo(rig.view.state).unwrap();
                rig.view.updateState(
                  publishCanonicalProjection({
                    session: rig.session,
                    state: rig.view.state,
                    commit: typingUndo,
                  }).unwrap().state,
                );
                expect(rig.view.state.doc.textContent).toBe(text);
                expect(rig.session.canUndo).toBe(true);
                const compositionUndo = rig.session.prepareUndo(rig.view.state).unwrap();
                rig.view.updateState(
                  publishCanonicalProjection({
                    session: rig.session,
                    state: rig.view.state,
                    commit: compositionUndo,
                  }).unwrap().state,
                );
                expect(rig.session.document).toEqual(original);
                expect(rig.view.state.doc.textContent).toBe("alpha");
                expect(rig.session.canUndo).toBe(false);
              } finally {
                rig.destroy();
              }
            }
          }
        }),
        { numRuns: 16, id: "same-payload typing after native end is always a separate gesture" },
      );
    } finally {
      jest.useRealTimers();
    }
  },
  propertyTestTimeout(5_000),
);
