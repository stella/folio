import { expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import { panic } from "better-result";
import { assertProperty, propertyTestTimeout } from "../../../../../test/property-testing";
import type { TextFormatting } from "../../types/document";

import { Fragment, Slice, type Node as PMNode } from "prosemirror-model";
import { TextSelection } from "prosemirror-state";
import { createDocx } from "../../docx/rezip";
import { createEmptyDocument } from "../../utils/createDocument";
import {
  createHarnessState,
  HeadlessEditorView,
  parseShapeDocument,
  saveHarnessState,
} from "../../__tests__/editorHarness";
import { toggleBold } from "../commands/formatting";

setDefaultTimeout(propertyTestTimeout(60_000));

const sourceDocument = async (runProperties?: TextFormatting) => {
  const source = createEmptyDocument({ initialText: "LR" });
  source.package.document.content = [
    {
      type: "paragraph",
      paraId: "12345678",
      ...(runProperties === undefined ? {} : { formatting: { runProperties } }),
      content: [
        { type: "run", content: [{ type: "text", text: "L" }] },
        ...Array.from({ length: 3 }, () => ({
          type: "run" as const,
          formatting: { bold: true },
          content: [{ type: "footnoteRef" as const, id: 123 }],
        })),
        { type: "run", content: [{ type: "text", text: "R" }] },
      ],
    },
  ];
  source.package.footnotes = [
    {
      type: "footnote",
      id: 123,
      content: createEmptyDocument({ initialText: "Note" }).package.document.content,
    },
  ];
  return parseShapeDocument(new Uint8Array(await createDocx(source)));
};

const tokens = (doc: PMNode) => {
  const result: unknown[] = [];
  doc.descendants((node) => {
    if (node.type.name === "paragraph") result.push({ type: "paragraph" });
    if (!node.isText) return;
    const ref = node.marks.find((mark) => mark.type.name === "footnoteRef");
    for (const text of node.text ?? "") {
      result.push({
        text,
        ref: ref ? { id: ref.attrs.id, kind: ref.attrs.noteType } : null,
        bold: node.marks.some((mark) => mark.type.name === "bold"),
        italic: node.marks.some((mark) => mark.type.name === "italic"),
      });
    }
  });
  return result;
};

test("seed 666618706 keeps note formatting after split, format, split and paste", async () => {
  const base = await sourceDocument();
  const state = createHarnessState(base, "editing");
  const reference = state.doc.nodeAt(2);
  if (!reference) return panic("Missing note reference");
  const referenceSlice = new Slice(Fragment.from(reference), 0, 0);
  const view = new HeadlessEditorView(state);
  const actions = [
    { kind: "split", left: 10, right: 0 },
    { kind: "bold", left: 0, right: 0 },
    { kind: "split", left: 0, right: 0 },
    { kind: "pasteReference", left: 13, right: 1 },
  ] as const;

  for (const action of actions) {
    const positions: number[] = [];
    view.state.doc.descendants((node, position) => {
      if (!node.isTextblock) return true;
      for (let offset = 0; offset <= node.content.size; offset += 1) {
        positions.push(position + 1 + offset);
      }
      return false;
    });
    const left = positions.at(action.left % positions.length);
    const right = positions.at(action.right % positions.length);
    if (left === undefined || right === undefined) return panic("Missing selection endpoint");
    view.state = view.state.apply(
      view.state.tr.setSelection(
        TextSelection.create(view.state.doc, Math.min(left, right), Math.max(left, right)),
      ),
    );
    if (action.kind === "split") view.pressKey("Enter");
    else if (action.kind === "bold") toggleBold(view.state, view.dispatch);
    else view.paste(referenceSlice);

    const saved = await saveHarnessState(view.state, base);
    const reopened = createHarnessState(await parseShapeDocument(saved.bytes), "editing");
    const referenceBold = (doc: PMNode) => {
      const values: boolean[] = [];
      doc.descendants((node) => {
        if (node.marks.some(({ type }) => type.name === "footnoteRef"))
          values.push(node.marks.some(({ type }) => type.name === "bold"));
      });
      return new Set(values);
    };
    // Occurrence multiplicity is covered by the note-occurrence oracle; this
    // regression isolates the authored formatting carried by surviving units.
    expect(referenceBold(reopened.doc)).toEqual(referenceBold(view.state.doc));
  }
});

test("split paragraphs retain authored run defaults through typed and pasted content", async () => {
  await assertProperty(
    fc.asyncProperty(
      fc.record({
        runProperties: fc.record({ bold: fc.boolean(), italic: fc.boolean() }),
        paragraphProperties: fc.record({
          contextualSpacing: fc.boolean(),
          snapToGrid: fc.boolean(),
          spaceAfter: fc.integer({ min: 0, max: 720 }),
        }),
        endpoint: fc.constantFrom("start", "middle", "end"),
        action: fc.constantFrom("type", "pasteReference"),
      }),
      async ({ runProperties, paragraphProperties, endpoint, action }) => {
        const source = await sourceDocument(runProperties);
        const paragraph = source.package.document.content.at(0);
        if (paragraph?.type !== "paragraph") return panic("Missing source paragraph.");
        // One authored reference keeps this invariant independent of adjacent occurrence coalescing.
        paragraph.content.splice(2, 2);
        paragraph.formatting = { ...paragraph.formatting, ...paragraphProperties };
        const base = await parseShapeDocument(new Uint8Array(await createDocx(source)));
        const view = new HeadlessEditorView(createHarnessState(base, "editing"));
        const reference = view.state.doc.nodeAt(2);
        if (!reference?.isText) return panic("Missing source reference.");
        const position = {
          start: 1,
          middle: view.state.doc.content.size - 2,
          end: view.state.doc.content.size - 1,
        }[endpoint];
        view.state = view.state.apply(
          view.state.tr.setSelection(TextSelection.create(view.state.doc, position)),
        );
        expect(view.pressKey("Enter")).toBe(true);
        expect(view.state.selection.$from.parent.attrs._originalFormatting).toMatchObject({
          runProperties,
          ...paragraphProperties,
        });
        if (action === "type") view.typeText("x");
        else view.paste(new Slice(Fragment.from(reference), 0, 0));
        const saved = await saveHarnessState(view.state, base);
        const addedParagraph = saved.model.package.document.content.at(1);
        if (addedParagraph?.type !== "paragraph") return panic("Missing split paragraph.");
        expect(addedParagraph.formatting).toMatchObject(paragraphProperties);
        const reopenedDocument = await parseShapeDocument(saved.bytes);
        const reopenedParagraph = reopenedDocument.package.document.content.at(1);
        if (reopenedParagraph?.type !== "paragraph")
          return panic("Missing reopened split paragraph.");
        expect(reopenedParagraph.formatting).toMatchObject(paragraphProperties);
        const reopened = createHarnessState(reopenedDocument, "editing");
        expect(tokens(reopened.doc)).toEqual(tokens(view.state.doc));
      },
    ),
    {
      numRuns: 30,
      id: "split paragraphs retain authored run defaults through typed and pasted content",
    },
  );
});
