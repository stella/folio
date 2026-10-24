import { describe, expect, test } from "bun:test";
import { EditorState, TextSelection } from "prosemirror-state";

import { getCanonicalCommandIntents } from "../../canonicalCommands";
import { toggleBold } from "../../commands/formatting";
import { schema } from "../../schema";

const lockedStates = () =>
  ["contentLocked", "sdtContentLocked"].flatMap((lock) =>
    [
      schema.node("doc", null, [
        schema.node("paragraph", null, [
          schema.node("sdt", { lock }, [schema.text("locked")]),
        ]),
      ]),
      schema.node("doc", null, [
        schema.node("blockSdt", { lock }, [
          schema.node("paragraph", null, [schema.text("locked")]),
        ]),
      ]),
    ].flatMap((doc) =>
      [TextSelection.create(doc, 2), TextSelection.create(doc, 2, 7)].map(
        (selection) => EditorState.create({ schema, doc, selection }),
      ),
    ),
  );

describe("run formatting respects locked content controls", () => {
  test("blocks collapsed and ranged formatting in inline and block controls", () => {
    for (const state of lockedStates()) {
      const originalDoc = state.doc;
      let dispatched = false;

      expect(
        toggleBold(state, () => {
          dispatched = true;
        }),
      ).toBe(false);
      expect(dispatched).toBe(false);
      expect(state.doc).toBe(originalDoc);
    }
  });

  test("canonical formatting probes report no intent inside locked controls", () => {
    for (const state of lockedStates()) {
      const originalDoc = state.doc;
      const originalSelection = state.selection;
      expect(getCanonicalCommandIntents(toggleBold, state)).toEqual([]);
      expect(state.doc).toBe(originalDoc);
      expect(state.selection).toBe(originalSelection);
    }
  });

  test("leaves unlocked and container-only locks formattable", () => {
    for (const lock of [null, "sdtLocked"]) {
      const doc = schema.node("doc", null, [
        schema.node("paragraph", null, [
          schema.node("sdt", { lock }, [schema.text("editable")]),
        ]),
      ]);
      const state = EditorState.create({
        schema,
        doc,
        selection: TextSelection.create(doc, 2, 7),
      });
      expect(toggleBold(state)).toBe(true);
      expect(getCanonicalCommandIntents(toggleBold, state)).toEqual([
        expect.objectContaining({ type: "formatRun" }),
      ]);
    }
  });
});
