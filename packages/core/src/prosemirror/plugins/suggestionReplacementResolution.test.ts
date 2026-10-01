import { expect, test } from "bun:test";
import { EditorState, TextSelection } from "prosemirror-state";
import { Fragment, Slice } from "prosemirror-model";
import { panic } from "better-result";

import { DOCUMENT_SHAPES } from "../../__tests__/documentShapes";
import {
  HeadlessEditorView,
  createHarnessState,
  parseShapeDocument,
  saveHarnessState,
  textblocks,
} from "../../__tests__/editorHarness";
import { FolioDocxReviewer } from "../../ai-edits/headless";
import { acceptAllChanges, rejectAllChanges } from "../commands/comments";
import { paragraphPropertiesSnapshot } from "../commands/propertyChangeScope";
import { schema } from "../schema";
import { createSuggestionModePlugin, suggestionModeKey } from "./suggestionMode";

const paragraph = (text: string) => schema.node("paragraph", null, schema.text(text));

// Exercise the real composition lifecycle, including the browser-owned native
// replacement between its start and deferred end handlers.
test.each(["beforeEnd", "nativeEndFlush"] as const)(
  "%s: IME replacement preserves open paragraph edges across every text range",
  async (commitTiming) => {
    const baseline = schema.node("doc", null, [
      schema.node("paragraph", { alignment: "center", indentLeft: 240 }, schema.text("First")),
      schema.node("paragraph", { alignment: "right", indentLeft: 480 }, schema.text("Second")),
    ]);
    const properties = (doc: typeof baseline) =>
      textblocks(doc).map(({ node }) => paragraphPropertiesSnapshot(node));
    const ranges = [];
    for (let from = 1; from <= 6; from++) {
      for (let to = 8; to <= 14; to++) ranges.push({ from, to });
    }
    for (const { from, to } of ranges) {
      const plugin = createSuggestionModePlugin(true, "Reviewer");
      const view = new HeadlessEditorView(EditorState.create({ doc: baseline, plugins: [plugin] }));
      view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, from, to)));
      plugin.props.handleDOMEvents?.["compositionstart"]?.call(
        plugin,
        view as never,
        new Event("compositionstart"),
      );
      const direct = view.state.tr.insertText("alpha").doc;
      if (commitTiming === "beforeEnd") view.dispatch(view.state.tr.insertText("alpha"));
      plugin.props.handleDOMEvents?.["compositionend"]?.call(
        plugin,
        view as never,
        new Event("compositionend"),
      );
      if (commitTiming === "nativeEndFlush") {
        queueMicrotask(() => view.dispatch(view.state.tr.insertText("alpha")));
      }
      await Promise.resolve();
      await Promise.resolve();
      let accepted = view.state;
      acceptAllChanges()(accepted, (tr) => {
        accepted = accepted.apply(tr);
      });
      expect(accepted.doc.textContent).toBe(direct.textContent);
      expect(accepted.doc.childCount).toBe(direct.childCount);
      expect(properties(accepted.doc)).toEqual(properties(direct));
      let rejected = view.state;
      rejectAllChanges()(rejected, (tr) => {
        rejected = rejected.apply(tr);
      });
      expect(rejected.doc.textContent).toBe(baseline.textContent);
      expect(rejected.doc.childCount).toBe(baseline.childCount);
      expect(properties(rejected.doc)).toEqual(properties(baseline));
    }
  },
);

const project = (reviewer: FolioDocxReviewer) =>
  reviewer.snapshot().blocks.map(({ kind, text, displayLabel, listLevel, table }) => ({
    kind,
    text,
    displayLabel,
    listLevel,
    table,
  }));

test.each(["mixed-lists", "image", "notes"])(
  "replacement traces preserve reader accept/reject projections in %s",
  async (shapeId) => {
    const shape = DOCUMENT_SHAPES.find(({ id }) => id === shapeId) ?? panic("Missing shape");
    const bytes = await shape.build();
    const base = await parseShapeDocument(bytes);
    const baseline = project(await FolioDocxReviewer.fromBuffer(bytes.slice().buffer));
    const outputs = [];
    for (const mode of ["editing", "suggesting"] as const) {
      const view = new HeadlessEditorView(createHarnessState(base, mode));
      if (shapeId === "notes") {
        view.pressKey("Delete");
        view.paste(
          new Slice(
            Fragment.from(
              schema.node("table", null, [
                schema.node("tableRow", null, [
                  schema.node("tableCell", null, [paragraph("Left")]),
                  schema.node("tableCell", null, [paragraph("Right")]),
                ]),
              ]),
            ),
            0,
            0,
          ),
        );
        view.paste(new Slice(Fragment.from(schema.text("alpha")), 0, 0));
      } else {
        const blocks = textblocks(view.state.doc);
        const first = blocks.at(1) ?? panic("Missing first replacement paragraph");
        const last = blocks.at(2) ?? panic("Missing last replacement paragraph");
        view.dispatch(
          view.state.tr.setSelection(
            TextSelection.create(view.state.doc, first.pos + 1, last.pos + 2),
          ),
        );
        const plugin = suggestionModeKey.get(view.state) ?? panic("Missing suggestion plugin");
        plugin.props.handleDOMEvents?.["compositionstart"]?.call(
          plugin,
          view as never,
          new Event("compositionstart"),
        );
        view.dispatch(view.state.tr.insertText("alpha"));
        plugin.props.handleDOMEvents?.["compositionend"]?.call(
          plugin,
          view as never,
          new Event("compositionend"),
        );
        await Promise.resolve();
        await Promise.resolve();
        if (shapeId === "mixed-lists") view.typeText("alpha");
      }
      outputs.push(await saveHarnessState(view.state, base));
    }
    const edited = outputs.at(0) ?? panic("Missing direct output");
    const suggested = outputs.at(1) ?? panic("Missing tracked output");
    const accepted = await FolioDocxReviewer.fromBuffer(suggested.bytes.slice().buffer);
    accepted.acceptAll();
    expect(project(await FolioDocxReviewer.fromBuffer(await accepted.toBuffer()))).toEqual(
      project(await FolioDocxReviewer.fromBuffer(edited.bytes.slice().buffer)),
    );
    const rejected = await FolioDocxReviewer.fromBuffer(suggested.bytes.slice().buffer);
    rejected.rejectAll();
    expect(project(await FolioDocxReviewer.fromBuffer(await rejected.toBuffer()))).toEqual(
      baseline,
    );
  },
);

test.each([1, 2, 3, 4, 5])(
  "table paste after %s tracked deletions removes the same leading break as direct paste",
  (count) => {
    const baseline = schema.node("doc", null, [paragraph("Intro"), paragraph("Tail")]);
    const pastedTable = schema.node("table", null, [
      schema.node("tableRow", null, [
        schema.node("tableCell", null, [paragraph("Left")]),
        schema.node("tableCell", null, [paragraph("Right")]),
      ]),
    ]);
    const views = [false, true].map(
      (active) =>
        new HeadlessEditorView(
          EditorState.create({
            doc: baseline,
            plugins: [createSuggestionModePlugin(active, "Reviewer")],
          }),
        ),
    );
    for (const view of views) {
      for (let index = 0; index < count; index++) view.pressKey("Delete");
      view.paste(new Slice(Fragment.from(pastedTable), 0, 0));
      view.paste(new Slice(Fragment.from(schema.text("alpha")), 0, 0));
    }
    const edited = views.at(0)!;
    const suggested = views.at(1)!;
    let accepted = suggested.state;
    acceptAllChanges()(accepted, (tr) => {
      accepted = accepted.apply(tr);
    });
    const blocks = (state: EditorState) => {
      const result: { kind: string; text: string }[] = [];
      state.doc.forEach((node) => result.push({ kind: node.type.name, text: node.textContent }));
      return result;
    };
    expect(blocks(accepted)).toEqual(blocks(edited.state));
    let rejected = suggested.state;
    rejectAllChanges()(rejected, (tr) => {
      rejected = rejected.apply(tr);
    });
    expect(blocks(rejected)).toEqual(blocks(EditorState.create({ doc: baseline })));
  },
);

// Paste must not take ownership of a revision already attached to the paragraph.
test.each(["del", "ins", "moveFrom", "moveTo"])(
  "table paste preserves an existing %s paragraph revision",
  (kind) => {
    const info = { id: 987, author: "Previous reviewer", date: "2026-01-01T00:00:00Z" };
    const pPrMark = { kind, info };
    const baseline = schema.node("doc", null, [
      schema.node(
        "paragraph",
        { pPrMark },
        schema.text("Intro", [
          schema.mark("deletion", {
            revisionId: info.id,
            author: info.author,
            date: info.date,
          }),
        ]),
      ),
      paragraph("Tail"),
    ]);
    const view = new HeadlessEditorView(
      EditorState.create({
        doc: baseline,
        plugins: [createSuggestionModePlugin(true, "Reviewer")],
      }),
    );
    view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, 6)));
    const table = schema.node("table", null, [
      schema.node("tableRow", null, [schema.node("tableCell", null, [paragraph("Pasted")])]),
    ]);
    view.paste(new Slice(Fragment.from(table), 0, 0));
    const restored = textblocks(view.state.doc).find(({ node }) => node.textContent === "Intro");
    expect(restored?.node.attrs["pPrMark"]).toEqual(pPrMark);
  },
);

// The physical caret can sit anywhere in struck text at the visible start.
// Exercise both deleted and live suffixes, including the final-paragraph guard.
test.each(["followed", "final"])(
  "table paste projects every caret in a deleted prefix of a %s paragraph",
  (placement) => {
    const baseline = schema.node("doc", null, [
      paragraph("Intro"),
      ...(placement === "followed" ? [paragraph("Tail")] : []),
    ]);
    const table = schema.node("table", null, [
      schema.node("tableRow", null, [schema.node("tableCell", null, [paragraph("Pasted")])]),
    ]);
    for (let count = 1; count <= 5; count++) {
      for (let caret = 2; caret <= count + 1; caret++) {
        const views = [false, true].map(
          (active) =>
            new HeadlessEditorView(
              EditorState.create({
                doc: baseline,
                plugins: [createSuggestionModePlugin(active, "Reviewer")],
              }),
            ),
        );
        for (const view of views) {
          for (let index = 0; index < count; index++) view.pressKey("Delete");
          if (suggestionModeKey.getState(view.state)?.active) {
            view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, caret)));
          }
          view.paste(new Slice(Fragment.from(table), 0, 0));
        }
        const edited = views.at(0) ?? panic("Missing edited state");
        const suggested = views.at(1) ?? panic("Missing suggested state");
        let accepted = suggested.state;
        acceptAllChanges()(accepted, (tr) => {
          accepted = accepted.apply(tr);
        });
        // A container must retain its final paragraph even when it is empty.
        const expected =
          placement === "final" && count === 5
            ? schema.node("doc", null, [table, schema.node("paragraph")])
            : edited.state.doc;
        expect(accepted.doc.toJSON()).toEqual(expected.toJSON());
        let rejected = suggested.state;
        rejectAllChanges()(rejected, (tr) => {
          rejected = rejected.apply(tr);
        });
        expect(rejected.doc.toJSON()).toEqual(baseline.toJSON());
      }
    }
  },
);
