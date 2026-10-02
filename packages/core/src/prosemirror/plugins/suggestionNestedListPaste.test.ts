import { afterAll, beforeAll, expect, test } from "bun:test";
import { GlobalRegistrator } from "@happy-dom/global-registrator";
import fc from "fast-check";
import { DOMParser } from "prosemirror-model";
import { EditorState, TextSelection } from "prosemirror-state";

import { assertProperty, propertyTestTimeout } from "../../../../../test/property-testing";
import { shapeArrayBuffer } from "../../__tests__/documentShapes";
import {
  createHarnessPlugins,
  createHarnessState,
  HeadlessEditorView,
  parseShapeDocument,
  resolveAllChanges,
  saveHarnessState,
  summarizeState,
} from "../../__tests__/editorHarness";
import { FolioDocxReviewer } from "../../ai-edits/headless";
import { expectParagraphAttrs } from "../attrs";
import { flattenPastedHtmlLists } from "../extensions/features/pastedHtmlLists";
import type { ParagraphAttrs } from "../schema/nodes";

beforeAll(() => GlobalRegistrator.register());
afterAll(() => GlobalRegistrator.unregister());

const nestedListHtml = (kind: "ul" | "ol") =>
  `<${kind}><li>Alpha<${kind}><li>Nested</li></${kind}></li><li>Omega</li></${kind}>`;

type PasteAndJoinOptions = {
  state: EditorState;
  caret: number;
  html: string;
  direction: "backward" | "forward";
};

const pasteAndJoin = ({ state, caret, html, direction }: PasteAndJoinOptions) => {
  const view = new HeadlessEditorView(
    state.apply(state.tr.setSelection(TextSelection.create(state.doc, caret))),
  );
  const host = document.createElement("div");
  host.innerHTML = flattenPastedHtmlLists(html);
  view.paste(DOMParser.fromSchema(state.schema).parseSlice(host));
  if (direction === "forward") {
    view.dispatch(
      view.state.tr.setSelection(
        TextSelection.create(view.state.doc, view.state.selection.from - 2),
      ),
    );
  }
  view.pressKey(direction === "backward" ? "Backspace" : "Delete");
  return view.state;
};

// Clipboard tests covered paste alone; also retract an inserted break while
// retaining an existing paragraph's formatting.
test(
  "retracting pasted list breaks preserves acceptance and rejection for paragraph properties",
  async () => {
    const base = await parseShapeDocument(new Uint8Array(await shapeArrayBuffer("mixed-lists")));
    const initial = createHarnessState(base, "editing");
    const initialAttrs = expectParagraphAttrs(initial.doc.child(0));
    await assertProperty(
      fc.property(
        fc.record({
          caret: fc.integer({ min: 1, max: 17 }),
          kind: fc.constantFrom("ul", "ol"),
          direction: fc.constantFrom("backward", "forward"),
          alignment: fc.constantFrom("left", "center", "right", "both"),
          indentLeft: fc.integer({ min: 0, max: 1000 }),
          indentFirstLine: fc.integer({ min: 0, max: 500 }),
          hangingIndent: fc.boolean(),
          spaceAfter: fc.integer({ min: 0, max: 500 }),
        }),
        ({
          caret,
          kind,
          direction,
          alignment,
          indentLeft,
          indentFirstLine,
          hangingIndent,
          spaceAfter,
        }) => {
          const paragraphAttrs = {
            ...initialAttrs,
            alignment,
            indentLeft,
            indentFirstLine,
            hangingIndent,
            spaceAfter,
            _originalFormatting: {
              ...initialAttrs._originalFormatting,
              alignment,
              indentLeft,
              indentFirstLine: hangingIndent ? -indentFirstLine : indentFirstLine,
              hangingIndent,
              spaceAfter,
            },
          } as const satisfies ParagraphAttrs;
          const before = initial.apply(initial.tr.setNodeMarkup(0, undefined, paragraphAttrs));
          const options = { caret, html: nestedListHtml(kind), direction };
          const edited = pasteAndJoin({ state: before, ...options });
          const tracked = pasteAndJoin({
            state: EditorState.create({
              doc: before.doc,
              plugins: createHarnessPlugins(base, "suggesting"),
            }),
            ...options,
          });
          expect(summarizeState(resolveAllChanges(tracked, "accept"))).toEqual(
            summarizeState(edited),
          );
          expect(summarizeState(resolveAllChanges(tracked, "reject"))).toEqual(
            summarizeState(before),
          );
        },
      ),
      { seed: 197, numRuns: 30 },
    );
  },
  propertyTestTimeout(30_000),
);

test("nested list paste and backspace rejects to the original paragraph after DOCX reopen", async () => {
  const base = await parseShapeDocument(new Uint8Array(await shapeArrayBuffer("mixed-lists")));
  const baseline = createHarnessState(base, "editing");
  const options = { caret: 1, html: nestedListHtml("ul"), direction: "backward" } as const;
  const edited = pasteAndJoin({ state: baseline, ...options });
  const tracked = pasteAndJoin({ state: createHarnessState(base, "suggesting"), ...options });
  const { bytes } = await saveHarnessState(tracked, base);
  const accepting = await FolioDocxReviewer.fromBuffer(bytes.slice().buffer);
  accepting.acceptAll();
  const accepted = createHarnessState(
    await parseShapeDocument(new Uint8Array(await accepting.toBuffer())),
    "editing",
  );
  const rejecting = await FolioDocxReviewer.fromBuffer(bytes.slice().buffer);
  rejecting.rejectAll();
  const rejected = createHarnessState(
    await parseShapeDocument(new Uint8Array(await rejecting.toBuffer())),
    "editing",
  );
  expect(summarizeState(accepted)).toEqual(summarizeState(edited));
  expect(summarizeState(rejected)).toEqual(summarizeState(baseline));
});

test("retracting a pasted break preserves the following paragraph's independent closing revision", async () => {
  const base = await parseShapeDocument(new Uint8Array(await shapeArrayBuffer("mixed-lists")));
  const initial = createHarnessState(base, "suggesting");
  const originalMark = {
    kind: "del",
    info: { id: 901, author: "Previous", date: "2026-01-01T00:00:00Z" },
  } as const;
  const before = initial.apply(initial.tr.setNodeAttribute(0, "pPrMark", originalMark));
  const after = pasteAndJoin({
    state: before,
    caret: 1,
    html: nestedListHtml("ul"),
    direction: "backward",
  });
  const joined = after.doc.child(2);
  expect(joined.textContent).toBe("OmegaIntro paragraph.");
  expect(expectParagraphAttrs(joined).pPrMark).toEqual(originalMark);
});
