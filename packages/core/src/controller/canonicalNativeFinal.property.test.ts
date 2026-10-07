import { afterAll, beforeAll, expect, test } from "bun:test";
import { GlobalRegistrator } from "@happy-dom/global-registrator";
import fc from "fast-check";
import { EditorState, TextSelection } from "prosemirror-state";
import { EditorView } from "prosemirror-view";
import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";
import { schema } from "../prosemirror/schema";
import { createEmptyDocument } from "../utils/createDocument";
import { createCanonicalInputBoundary } from "./canonicalInput";
import { createCanonicalSession, publishCanonicalProjection } from "./canonicalSession";

beforeAll(() => GlobalRegistrator.register());
afterAll(() => GlobalRegistrator.unregister());

const createRig = () => {
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
    state: EditorState.create({ schema, doc, selection: TextSelection.create(doc, 1, 6) }),
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
    assertProperty(
      fc.property(
        fc.constantFrom("契", "😀", "مرحبا", "alpha", "alphabet", "alpine"),
        fc.constantFrom("provisional", "refused", "empty"),
        (text, phase) => {
          const rig = createRig();
          try {
            const document = rig.session.document;
            const baseline = rig.view.state;
            rig.view.dom.dispatchEvent(new Event("compositionstart", { bubbles: true }));
            if (phase !== "empty") {
              const provisional = rig.view.state.tr
                .insertText(text, 1, 6)
                .setMeta("composition", 1);
              if (phase === "refused") provisional.addMark(1, 1 + text.length, schema.mark("bold"));
              rig.view.dispatch(provisional);
            }
            expect(rig.view.composing).toBe(true);
            const final = new InputEvent("beforeinput", {
              bubbles: true,
              cancelable: true,
              inputType: "insertText",
              data: text,
            });
            // No compositionend or cleanup cancel precedes this final input.
            rig.view.dom.dispatchEvent(final);
            expect(final.defaultPrevented).toBe(true);
            expect(rig.view.composing).toBe(false);
            expect(rig.boundary.isComposing).toBe(false);
            expect(rig.refusals).toHaveLength(phase === "refused" ? 1 : 0);
            expect(rig.view.state.doc.textContent).toBe(phase === "refused" ? "alpha" : text);
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
        },
      ),
      { numRuns: 36 },
    );
  },
  propertyTestTimeout(5_000),
);

test("a different typing payload after recovery remains a separate edit", () => {
  const rig = createRig();
  try {
    rig.view.dom.dispatchEvent(new Event("compositionstart", { bubbles: true }));
    rig.view.dispatch(rig.view.state.tr.insertText("契", 1, 6).setMeta("composition", 1));
    rig.view.dom.dispatchEvent(
      new InputEvent("beforeinput", {
        bubbles: true,
        cancelable: true,
        inputType: "insertText",
        data: "NEXT",
      }),
    );
    expect(rig.refusals).toEqual([]);
    expect(rig.view.state.doc.textContent).toBe("契NEXT");
    expect(rig.session.version).toBe(2);
    const undo = rig.session.prepareUndo(rig.view.state).unwrap();
    rig.view.updateState(
      publishCanonicalProjection({
        session: rig.session,
        state: rig.view.state,
        commit: undo,
      }).unwrap().state,
    );
    expect(rig.view.state.doc.textContent).toBe("契");
    expect(rig.session.canUndo).toBe(true);
  } finally {
    rig.destroy();
  }
});
