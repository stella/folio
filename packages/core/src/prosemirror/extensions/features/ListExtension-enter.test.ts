import { describe, expect, test } from "bun:test";
import { panic } from "better-result";
import { EditorState, TextSelection } from "prosemirror-state";

import { toFlowBlocks } from "../../../layout-bridge/convert/toFlowBlocks";
import type { Document } from "../../../types/document";
import { toProseDoc } from "../../conversion/toProseDoc";
import { LIST_RENDERING_ATTR_KEYS } from "../../listMarker";
import { schema } from "../../schema";
import { ListExtension, toggleNumberedList } from "./ListExtension";

const syntheticNumberedDocument = (): Document => ({
  package: {
    document: {
      content: [
        {
          type: "paragraph",
          formatting: { numPr: { numId: 23, ilvl: 3 } },
          content: [
            {
              type: "run",
              content: [{ type: "text", text: "Synthetic numbered clause" }],
            },
          ],
          listRendering: {
            marker: "(a)",
            markerTemplate: "(%4)",
            level: 3,
            numId: 23,
            isBullet: false,
            numFmt: "lowerLetter",
            levelNumFmts: ["decimal", "decimal", "decimal", "lowerLetter"],
            abstractNumId: 15,
            startOverride: 1,
          },
        },
      ],
    },
  },
});

const listMarkers = (state: EditorState): string[] =>
  toFlowBlocks(state.doc).flatMap((block) =>
    block.kind === "paragraph" && block.attrs?.listMarker ? [block.attrs.listMarker] : [],
  );

describe("ListExtension Enter numbering", () => {
  test("advances an imported marker from its source template", () => {
    let state = EditorState.create({ doc: toProseDoc(syntheticNumberedDocument()) });
    const paragraph = state.doc.firstChild;
    if (!paragraph) {
      panic("Synthetic document did not contain its paragraph");
    }
    state = state.apply(
      state.tr.setSelection(TextSelection.create(state.doc, paragraph.nodeSize - 1)),
    );

    const enter = ListExtension().onSchemaReady({ schema }).keyboardShortcuts?.["Enter"];
    expect(enter).toBeDefined();
    if (!enter) {
      panic("List extension did not register Enter");
    }
    expect(
      enter(state, (transaction) => {
        state = state.apply(transaction);
      }),
    ).toBe(true);

    expect(listMarkers(state)).toEqual(["(a)", "(b)"]);
  });

  test("does not retain an imported template after replacing the list", () => {
    let state = EditorState.create({ doc: toProseDoc(syntheticNumberedDocument()) });
    const runtime = ListExtension().onSchemaReady({ schema });
    const removeList = runtime.commands?.removeList?.();
    if (!removeList) {
      panic("List extension did not register removeList");
    }

    expect(
      removeList(state, (transaction) => {
        state = state.apply(transaction);
      }),
    ).toBe(true);

    const attrs = state.doc.firstChild?.attrs;
    if (!attrs) {
      panic("Synthetic document did not contain its paragraph");
    }
    for (const key of LIST_RENDERING_ATTR_KEYS) {
      expect(attrs[key]).toBeNull();
    }

    expect(
      toggleNumberedList(state, (transaction) => {
        state = state.apply(transaction);
      }),
    ).toBe(true);

    const paragraph = state.doc.firstChild;
    if (!paragraph) {
      panic("Synthetic document did not contain its paragraph");
    }
    state = state.apply(
      state.tr.setSelection(TextSelection.create(state.doc, paragraph.nodeSize - 1)),
    );

    const enter = runtime.keyboardShortcuts?.["Enter"];
    if (!enter) {
      panic("List extension did not register Enter");
    }
    expect(
      enter(state, (transaction) => {
        state = state.apply(transaction);
      }),
    ).toBe(true);

    expect(listMarkers(state)).toEqual(["1.", "2."]);
  });

  test("clears every rendering attr when decreasing the first level exits the list", () => {
    let state = EditorState.create({
      doc: toProseDoc({
        ...syntheticNumberedDocument(),
        package: {
          document: {
            content: [
              {
                ...syntheticNumberedDocument().package.document.content[0],
                formatting: { numPr: { numId: 23, ilvl: 0 } },
              },
            ],
          },
        },
      }),
    });
    const decreaseListLevel = ListExtension()
      .onSchemaReady({ schema })
      .commands?.decreaseListLevel?.();
    if (!decreaseListLevel) {
      panic("List extension did not register decreaseListLevel");
    }

    expect(
      decreaseListLevel(state, (transaction) => {
        state = state.apply(transaction);
      }),
    ).toBe(true);

    const attrs = state.doc.firstChild?.attrs;
    if (!attrs) {
      panic("Synthetic document did not contain its paragraph");
    }
    for (const key of LIST_RENDERING_ATTR_KEYS) {
      expect(attrs[key]).toBeNull();
    }
  });
});
