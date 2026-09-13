import { describe, expect, test } from "bun:test";
import { EditorState } from "prosemirror-state";

import { schema } from "../schema";
import {
  createDocumentStylesPlugin,
  documentStylesKey,
  getDocumentStyleResolver,
  withDocumentStyles,
} from "./documentStyles";

describe("document style resolver state", () => {
  test("replaces the resolver without duplicating its plugin", () => {
    const state = EditorState.create({
      schema,
      plugins: [
        createDocumentStylesPlugin({
          styles: [{ styleId: "Normal", type: "paragraph", default: true }],
        }),
      ],
    });
    const styles = {
      styles: [
        { styleId: "Normal", type: "paragraph", default: true },
        {
          styleId: "ImportedCharacter",
          type: "character",
          rPr: { underline: { style: "double" } },
        },
      ],
    };

    const refreshed = withDocumentStyles(state, styles);

    expect(refreshed.doc.eq(state.doc)).toBe(true);
    expect(
      refreshed.plugins.filter((plugin) => plugin === documentStylesKey.get(refreshed)),
    ).toHaveLength(1);
    expect(getDocumentStyleResolver(state)?.getStyle("ImportedCharacter")).toBeUndefined();
    expect(getDocumentStyleResolver(refreshed)?.getStyle("ImportedCharacter")?.rPr).toEqual({
      underline: { style: "double" },
    });
  });
});
