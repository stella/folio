import { expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import { panic } from "better-result";
import type { Node as PMNode } from "prosemirror-model";
import { TextSelection } from "prosemirror-state";
import { assertProperty, propertyTestTimeout } from "../../../../../../test/property-testing";
import { createEmptyDocument } from "../../../utils/createDocument";
import { createDocx } from "../../../docx/rezip";
import {
  createHarnessState,
  HeadlessEditorView,
  parseShapeDocument,
  saveHarnessState,
} from "../../../__tests__/editorHarness";
import { toggleBold, toggleItalic } from "../../commands/formatting";
import { singletonManager } from "../../schema";
import { clearFormatting } from "./markUtils";

setDefaultTimeout(propertyTestTimeout(60_000));

const sourceDocument = async () => {
  const source = createEmptyDocument({ initialText: "" });
  source.package.document.content = [
    {
      type: "paragraph",
      paraId: "12345678",
      formatting: { alignment: "right", keepNext: true, spaceAfter: 120 },
      content: [],
    },
  ];
  return parseShapeDocument(new Uint8Array(await createDocx(source)));
};

const tokens = (doc: PMNode) => {
  const result: unknown[] = [];
  doc.descendants((node) => {
    if (node.isTextblock) result.push({ paragraph: true });
    if (!node.isText) return;
    const revision = node.marks.find(
      ({ type }) => type.name === "insertion" || type.name === "deletion",
    );
    for (const text of node.text ?? "")
      result.push({
        text,
        bold: node.marks.some(({ type }) => type.name === "bold"),
        italic: node.marks.some(({ type }) => type.name === "italic"),
        revision: revision
          ? {
              type: revision.type.name,
              author: revision.attrs.author,
              date: revision.attrs.date ? new Date(revision.attrs.date).toISOString() : null,
            }
          : null,
      });
  });
  return result;
};

const assertRoundtrip = async (
  view: HeadlessEditorView,
  base: Awaited<ReturnType<typeof sourceDocument>>,
) => {
  const saved = await saveHarnessState(view.state, base);
  const reopened = createHarnessState(await parseShapeDocument(saved.bytes), "editing");
  expect(tokens(reopened.doc)).toEqual(tokens(view.state.doc));
  const paragraph = saved.model.package.document.content.at(0);
  if (paragraph?.type !== "paragraph") return panic("Missing source paragraph.");
  expect(paragraph.formatting).toMatchObject({
    alignment: "right",
    keepNext: true,
    spaceAfter: 120,
  });
  return saved;
};

// ProseMirror's doPaste uses replaceSelectionWith(singleNode, true) for plain text.
const paste = (view: HeadlessEditorView) =>
  view.dispatch(
    view.state.tr
      .replaceSelectionWith(view.state.schema.text("x"), true)
      .setMeta("paste", true)
      .setMeta("uiEvent", "paste"),
  );

test("empty paragraph formatting and explicit clearing survive typing, paste, save and reopen", async () => {
  const base = await sourceDocument();
  for (const mode of ["editing", "suggesting"] as const) {
    for (const input of ["type", "paste"] as const) {
      for (const clear of [false, true]) {
        const view = new HeadlessEditorView(createHarnessState(base, mode));
        expect(toggleBold(view.state, view.dispatch)).toBe(true);
        expect(view.state.doc.firstChild?.attrs._originalFormatting?.runProperties?.bold).toBe(
          true,
        );
        if (clear) {
          expect(clearFormatting(view.state, view.dispatch)).toBe(true);
          expect(
            view.state.doc.firstChild?.attrs._originalFormatting?.runProperties,
          ).toBeUndefined();
        }
        if (input === "type") view.typeText("x");
        else paste(view);
        const text = view.state.doc.firstChild?.firstChild;
        expect(text?.marks.some(({ type }) => type.name === "bold")).toBe(!clear);
        await assertRoundtrip(view, base);
      }
    }
  }
});

test("random formatting and text edit sequences preserve authored paragraph defaults through save and reopen", async () => {
  const base = await sourceDocument();
  const editAction = fc.record({
    kind: fc.constantFrom(
      "bold",
      "italic",
      "clear",
      "type",
      "paste",
      "delete",
      "split",
      "join",
      "undo",
      "redo",
    ),
    left: fc.nat(20),
    right: fc.nat(20),
  });
  await assertProperty(
    fc.asyncProperty(
      fc.constantFrom("editing", "suggesting"),
      fc.array(editAction, { minLength: 1, maxLength: 12 }),
      async (mode, actions) => {
        const view = new HeadlessEditorView(createHarnessState(base, mode));
        for (const action of actions) {
          const positions: number[] = [];
          view.state.doc.descendants((node, position) => {
            if (!node.isTextblock) return true;
            for (let offset = 0; offset <= node.content.size; offset += 1)
              positions.push(position + offset + 1);
            return false;
          });
          const left = positions.at(action.left % positions.length);
          const right = positions.at(action.right % positions.length);
          if (left === undefined || right === undefined) return panic("Missing edit endpoint.");
          view.state = view.state.apply(
            view.state.tr.setSelection(
              TextSelection.create(view.state.doc, Math.min(left, right), Math.max(left, right)),
            ),
          );
          switch (action.kind) {
            case "bold":
              toggleBold(view.state, view.dispatch);
              break;
            case "italic":
              toggleItalic(view.state, view.dispatch);
              break;
            case "clear":
              clearFormatting(view.state, view.dispatch);
              break;
            case "type":
              view.typeText("x");
              break;
            case "paste":
              paste(view);
              break;
            case "delete":
              view.pressKey("Delete");
              break;
            case "split":
              view.pressKey("Enter");
              break;
            case "join":
              view.pressKey("Backspace");
              break;
            case "undo":
              singletonManager.requireCommand("undo")()(view.state, view.dispatch);
              break;
            case "redo":
              singletonManager.requireCommand("redo")()(view.state, view.dispatch);
              break;
            default:
              panic(String(action.kind satisfies never));
          }
          await assertRoundtrip(view, base);
        }
      },
    ),
    { numRuns: 20 },
  );
});
