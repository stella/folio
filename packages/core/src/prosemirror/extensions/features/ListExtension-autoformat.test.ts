import { describe, expect, test } from "bun:test";
import { undoInputRule } from "prosemirror-inputrules";
import { EditorState, TextSelection } from "prosemirror-state";
import type { Plugin, Transaction } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";

import { schema } from "../../schema";
import { ListExtension } from "./ListExtension";

const autoformatPlugin = (): Plugin => {
  const plugins = ListExtension().onSchemaReady({ schema }).plugins ?? [];
  const plugin = plugins.at(0);
  if (!plugin) throw new Error("the list extension registered no plugin");
  return plugin;
};

type TypedParagraph = {
  text: string;
  attrs: Record<string, unknown>;
  state: EditorState;
};

/**
 * A paragraph holding `marker`, with the space that triggers the rule fed
 * through `handleTextInput` the way the editor's text-input funnel feeds it.
 * A space no rule claims is inserted as the funnel's default would insert it.
 */
const typeMarker = (
  marker: string,
  { leading = "", listAttrs = {} }: { leading?: string; listAttrs?: Record<string, unknown> } = {},
): TypedParagraph => {
  const plugin = autoformatPlugin();
  const content = `${leading}${marker}`;
  const view = {
    composing: false,
    state: EditorState.create({
      doc: schema.node("doc", null, [
        schema.node("paragraph", listAttrs, content === "" ? undefined : [schema.text(content)]),
      ]),
      plugins: [plugin],
    }),
    dispatch(tr: Transaction) {
      view.state = view.state.apply(tr);
    },
  };
  const caret = content.length + 1;
  view.state = view.state.apply(
    view.state.tr.setSelection(TextSelection.create(view.state.doc, caret)),
  );

  const handled = plugin.props.handleTextInput?.(view as unknown as EditorView, caret, caret, " ");
  if (!handled) {
    view.state = view.state.apply(view.state.tr.insertText(" ", caret));
  }

  const paragraph = view.state.doc.child(0);
  return { text: paragraph.textContent, attrs: paragraph.attrs, state: view.state };
};

describe("list autoformat", () => {
  test.each(["-", "*"])("%s followed by a space starts a bulleted list", (marker) => {
    const { text, attrs } = typeMarker(marker);

    expect(text).toBe("");
    expect(attrs["numPr"]).toEqual({ numId: 1, ilvl: 0 });
    expect(attrs["listIsBullet"]).toBe(true);
  });

  test("1. followed by a space starts a numbered list", () => {
    const { text, attrs } = typeMarker("1.");

    expect(text).toBe("");
    expect(attrs["numPr"]).toEqual({ numId: 2, ilvl: 0 });
    expect(attrs["listIsBullet"]).toBe(false);
    expect(attrs["listNumFmt"]).toBe("decimal");
  });

  test("backspace right after the conversion puts the typed marker back", () => {
    const { state } = typeMarker("-");
    let undone = state;
    const handled = undoInputRule(state, (tr) => {
      undone = state.apply(tr);
    });

    expect(handled).toBe(true);
    expect(undone.doc.child(0).textContent).toBe("- ");
    expect(undone.doc.child(0).attrs["numPr"]).toBeNull();
  });

  test("a marker typed mid-paragraph stays text", () => {
    const { text, attrs } = typeMarker("-", { leading: "Item " });

    expect(text).toBe("Item - ");
    expect(attrs["numPr"]).toBeNull();
  });

  test("a marker typed in a list item does not toggle the list off", () => {
    const { text, attrs } = typeMarker("-", {
      listAttrs: { numPr: { numId: 1, ilvl: 0 }, listIsBullet: true },
    });

    expect(text).toBe("- ");
    expect(attrs["numPr"]).toEqual({ numId: 1, ilvl: 0 });
  });

  test("a number other than one stays text", () => {
    const { text, attrs } = typeMarker("2.");

    expect(text).toBe("2. ");
    expect(attrs["numPr"]).toBeNull();
  });
});
