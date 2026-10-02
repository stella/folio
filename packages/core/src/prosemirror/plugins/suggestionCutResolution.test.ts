import { expect, test } from "bun:test";
import { panic } from "better-result";
import { Fragment, Slice } from "prosemirror-model";
import { NodeSelection, TextSelection } from "prosemirror-state";

import { documentShape } from "../../__tests__/documentShapes";
import {
  createHarnessState,
  HeadlessEditorView,
  parseShapeDocument,
  saveHarnessState,
} from "../../__tests__/editorHarness";
import { FolioDocxReviewer } from "../../ai-edits/headless";
import { deleteSelectionAsSuggestion } from "./suggestionMode";

const project = (reviewer: FolioDocxReviewer) =>
  reviewer.snapshot().blocks.map(({ kind, text, displayLabel, listLevel, table }) => ({
    kind,
    text,
    displayLabel,
    listLevel,
    table,
  }));

const cases = (["open", "closed"] as const).flatMap((clipboard) =>
  (["node", "image-range", "paragraph-range", "cross-paragraph"] as const).flatMap((range) =>
    (["forward", "backward"] as const).map((direction) => ({ clipboard, range, direction })),
  ),
);

test.each(cases)(
  "$clipboard paste then $direction $range cut resolves equally",
  async ({ clipboard, range, direction }) => {
    const bytes = await documentShape("image").build();
    const base = await parseShapeDocument(bytes);
    const baseline = project(await FolioDocxReviewer.fromBuffer(bytes.slice().buffer));
    const outputs = [];
    for (const mode of ["editing", "suggesting"] as const) {
      const view = new HeadlessEditorView(createHarnessState(base, mode));
      const { schema } = view.state;
      const strong = schema.mark("bold");
      view.paste(
        new Slice(
          Fragment.from([
            schema.node("paragraph", null, [schema.text("First "), schema.text("bold", [strong])]),
            schema.node("paragraph", null, [schema.text("Second")]),
          ]),
          clipboard === "open" ? 1 : 0,
          clipboard === "open" ? 1 : 0,
        ),
      );
      let image = -1;
      let tail = -1;
      view.state.doc.descendants((node, pos) => {
        if (node.type.name === "image") image = pos;
        if (node.isTextblock && node.textContent === "Tail.") tail = pos + 1;
      });
      if (image < 0 || tail < 0) panic("Missing selection fixture positions");
      const paragraphStart = view.state.doc.resolve(image).start();
      const from = range === "paragraph-range" ? paragraphStart : image;
      const to = range === "cross-paragraph" ? tail + 2 : image + 3;
      const selection =
        range === "node"
          ? NodeSelection.create(view.state.doc, image)
          : TextSelection.create(
              view.state.doc,
              direction === "forward" ? from : to,
              direction === "forward" ? to : from,
            );
      view.dispatch(view.state.tr.setSelection(selection));
      if (mode === "suggesting") {
        expect(deleteSelectionAsSuggestion(view.state, view.dispatch)).toBe(true);
      } else {
        view.dispatch(view.state.tr.deleteSelection());
      }
      outputs.push(await saveHarnessState(view.state, base));
    }
    const direct = outputs.at(0) ?? panic("Missing edited output");
    const suggested = outputs.at(1) ?? panic("Missing suggested output");
    const edited = project(await FolioDocxReviewer.fromBuffer(direct.bytes.slice().buffer));
    const accepted = await FolioDocxReviewer.fromBuffer(suggested.bytes.slice().buffer);
    accepted.acceptAll();
    expect(project(await FolioDocxReviewer.fromBuffer(await accepted.toBuffer()))).toEqual(edited);
    const rejected = await FolioDocxReviewer.fromBuffer(suggested.bytes.slice().buffer);
    rejected.rejectAll();
    expect(project(await FolioDocxReviewer.fromBuffer(await rejected.toBuffer()))).toEqual(
      baseline,
    );
  },
);
