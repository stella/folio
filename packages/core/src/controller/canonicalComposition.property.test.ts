import { afterAll, beforeAll, expect, setDefaultTimeout, test } from "bun:test";
import { GlobalRegistrator } from "@happy-dom/global-registrator";
import fc from "fast-check";
import { Schema } from "prosemirror-model";
import { EditorState, TextSelection } from "prosemirror-state";
import { EditorView } from "prosemirror-view";
import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";
import { createCanonicalInputBoundary } from "./canonicalInput";

setDefaultTimeout(propertyTestTimeout(30_000));
beforeAll(() => GlobalRegistrator.register());
afterAll(() => GlobalRegistrator.unregister());

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
