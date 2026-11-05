import { describe, expect, test } from "bun:test";
import { TextSelection } from "prosemirror-state";

import { shapeArrayBuffer } from "../../../__tests__/documentShapes";
import {
  createHarnessState,
  HeadlessEditorView,
  parseShapeDocument,
  textblocks,
} from "../../../__tests__/editorHarness";

const openView = async () => {
  const state = createHarnessState(
    await parseShapeDocument(
      new Uint8Array(await shapeArrayBuffer("single-bullet-list")),
    ),
    "editing",
  );
  const empty = state.schema.nodes["paragraph"]!.create();
  const withEmpty = state.apply(state.tr.insert(0, empty));
  return new HeadlessEditorView(
    withEmpty.apply(
      withEmpty.tr.setSelection(TextSelection.create(withEmpty.doc, 1)),
    ),
  );
};

const direction = (view: HeadlessEditorView) =>
  textblocks(view.state.doc).at(0)?.node.attrs["direction"];

const domEvent = (
  view: HeadlessEditorView,
  name: "compositionstart" | "compositionend",
) => {
  view.someProp("handleDOMEvents", (events) => {
    (
      events as unknown as Record<
        string,
        ((view: unknown, event: unknown) => boolean) | undefined
      >
    )[name]?.(view, {});
    return false;
  });
};

const endComposition = async (view: HeadlessEditorView) => {
  domEvent(view, "compositionend");
  await Promise.resolve();
  await Promise.resolve();
};

describe("a paragraph's direction during a composition", () => {
  test("waits for the composition to end", async () => {
    const view = await openView();
    const before = direction(view);

    domEvent(view, "compositionstart");
    view.dispatch(
      view.state.tr.insertText("مرحبا", 1, 1).setMeta("composition", 1),
    );
    expect(direction(view)).toEqual(before);

    await endComposition(view);
    expect(direction(view)).toEqual({ source: "auto" });
  });

  test("waits for it through another writer's edit too", async () => {
    const view = await openView();
    const before = direction(view);

    domEvent(view, "compositionstart");
    view.dispatch(view.state.tr.insertText("مرحبا", 1, 1));
    expect(direction(view)).toEqual(before);

    await endComposition(view);
    expect(direction(view)).toEqual({ source: "auto" });
  });

  test("follows text a late flush brings after it has ended", async () => {
    const view = await openView();

    domEvent(view, "compositionstart");
    await endComposition(view);
    view.dispatch(
      view.state.tr.insertText("مرحبا", 1, 1).setMeta("composition", 1),
    );
    expect(direction(view)).toEqual({ source: "auto" });
  });

  test("still follows ordinary typing", async () => {
    const view = await openView();
    view.dispatch(view.state.tr.insertText("مرحبا", 1, 1));
    expect(direction(view)).toEqual({ source: "auto" });
  });
});
