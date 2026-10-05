import { afterAll, beforeAll, expect, setDefaultTimeout, test } from "bun:test";
import { GlobalRegistrator } from "@happy-dom/global-registrator";
import fc from "fast-check";
import { EditorState, TextSelection } from "prosemirror-state";
import type { Transaction } from "prosemirror-state";
import type { Node as ProseMirrorNode } from "prosemirror-model";
import { EditorView } from "prosemirror-view";
import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";
import { DOCUMENT_SHAPES, shapeArrayBuffer } from "../__tests__/documentShapes";
import { parseDocx } from "../docx/parser";
import { splitsSurrogatePair } from "../ai-edits/character-boundaries";
import { createCanonicalComposition } from "./canonicalComposition";
import { createCanonicalSession } from "./canonicalSession";

setDefaultTimeout(propertyTestTimeout(60_000));

beforeAll(() => GlobalRegistrator.register());
afterAll(() => GlobalRegistrator.unregister());

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
  );
  expect(docs.length).toBeGreaterThan(0);
  expect(executed.markLoss).toBeGreaterThan(0);
  expect(executed.textEdit).toBeGreaterThan(0);
});
