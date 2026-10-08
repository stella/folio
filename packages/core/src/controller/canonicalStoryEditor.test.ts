import { afterAll, afterEach, beforeAll, expect, test } from "bun:test";
import fc from "fast-check";
import { GlobalRegistrator } from "@happy-dom/global-registrator";
import { Schema } from "prosemirror-model";
import { AllSelection, EditorState, TextSelection } from "prosemirror-state";
import { EditorView } from "prosemirror-view";
import type { OpStory } from "@stll/docx-core/ops";
import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";
import { createCanonicalStoryEditor } from "./canonicalStoryEditor";

const schema = new Schema({
  nodes: {
    doc: { content: "paragraph+" },
    paragraph: { content: "text*", toDOM: () => ["p", 0] },
    text: {},
  },
});
const stories = [
  { kind: "footnote", id: 1 },
  { kind: "endnote", id: 1 },
  { kind: "header", rId: "rId1" },
  { kind: "footer", rId: "rId2" },
] as const satisfies readonly OpStory[];
const views: EditorView[] = [];

beforeAll(() => GlobalRegistrator.register());
afterEach(() => {
  for (const view of views.splice(0)) {
    const mount = view.dom.parentElement;
    view.destroy();
    mount?.remove();
  }
});
afterAll(() => GlobalRegistrator.unregister());

for (const story of stories) {
  test(
    `${story.kind} typing addresses the native caret before selection observation`,
    () => {
      assertProperty(
        fc.property(
          fc.integer({ min: 0, max: 13 }),
          fc.integer({ min: 0, max: 13 }),
          (modelOffset, nativeOffset) => {
            const doc = schema.node("doc", null, [
              schema.node("paragraph", null, schema.text("Footnote seed")),
            ]);
            const mount = document.createElement("div");
            document.body.append(mount);
            const canonical = createCanonicalStoryEditor({
              story,
              getView: () => view,
              getApi: () => null,
              enabled: () => true,
              onSelectionChange: () => {},
              onRefusal: (message) => {
                throw new TypeError(message);
              },
            });
            const view = new EditorView(mount, {
              state: EditorState.create({
                doc,
                selection: TextSelection.create(doc, modelOffset + 1),
              }),
              ...canonical.props,
              dispatchTransaction: (transaction) => canonical.dispatch(transaction),
            });
            views.push(view);
            const caret = view.domAtPos(nativeOffset + 1);
            const range = document.createRange();
            range.setStart(caret.node, caret.offset);
            range.collapse(true);
            const selection = document.getSelection();
            selection?.removeAllRanges();
            selection?.addRange(range);
            view.dom.dispatchEvent(new KeyboardEvent("keyup", { key: "End" }));
            const typeText = () => {
              const event = new InputEvent("beforeinput", {
                inputType: "insertText",
                data: "x",
                cancelable: true,
              });
              Object.defineProperty(event, "getTargetRanges", { value: () => [range] });
              return canonical.props.handleDOMEvents.beforeinput(view, event);
            };
            expect(typeText()).toBe(true);
            expect(view.state.selection.from).toBe(nativeOffset + 1);
            expect(view.state.selection.to).toBe(nativeOffset + 1);
            expect(view.state.doc).toBe(doc);
            // Select-all remains a model selection spanning block boundaries.
            view.dispatch(view.state.tr.setSelection(new AllSelection(doc)));
            expect(typeText()).toBe(true);
            expect(view.state.selection).toBeInstanceOf(AllSelection);
            view.destroy();
            mount.remove();
            views.pop();
          },
        ),
        { numRuns: 60, seed: 2392 },
      );
    },
    propertyTestTimeout(30_000),
  );
}
