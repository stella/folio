import { describe, expect, test } from "bun:test";
import { Slice } from "prosemirror-model";
import { EditorState } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";

import type { StyleDefinitions } from "../../../types/document";
import { createDocumentStylesPlugin } from "../../plugins/documentStyles";
import { schema } from "../../schema";
import { retargetPastedHeadingStyles } from "./pastedHeadingStyles";

/** A Czech Word's styles: the built-in heading under a localized id. */
const LOCALIZED: StyleDefinitions = {
  styles: [
    { styleId: "Normln", type: "paragraph", name: "Normal", default: true },
    { styleId: "Nadpis1", type: "paragraph", name: "heading 1", pPr: { outlineLevel: 0 } },
    { styleId: "Nadpis2", type: "paragraph", name: "heading 2", pPr: { outlineLevel: 1 } },
  ],
};

/**
 * The slice the `h1`–`h6` paste rule produces: the English built-in id plus
 * the outline level it stands for.
 */
const pastedHeading = (level: number, text: string): Slice =>
  new Slice(
    schema.node("doc", null, [
      schema.node("paragraph", { styleId: `Heading${level}`, outlineLevel: level - 1 }, [
        schema.text(text),
      ]),
    ]).content,
    0,
    0,
  );

const viewWith = (styles: StyleDefinitions | null): EditorView =>
  ({
    state: EditorState.create({ schema, plugins: [createDocumentStylesPlugin(styles)] }),
  }) as EditorView;

const styleIdsOf = (slice: Slice): (string | null)[] => {
  const ids: (string | null)[] = [];
  slice.content.descendants((node) => {
    if (node.type.name === "paragraph") {
      ids.push(node.attrs["styleId"] as string | null);
    }
  });
  return ids;
};

describe("retargetPastedHeadingStyles", () => {
  test("points a pasted heading at the document's own heading style", () => {
    const view = viewWith(LOCALIZED);
    expect(styleIdsOf(retargetPastedHeadingStyles(pastedHeading(1, "Smlouva"), view))).toEqual([
      "Nadpis1",
    ]);
    expect(styleIdsOf(retargetPastedHeadingStyles(pastedHeading(2, "Strany"), view))).toEqual([
      "Nadpis2",
    ]);
  });

  test("keeps the text and the outline level", () => {
    const retargeted = retargetPastedHeadingStyles(
      pastedHeading(1, "Smlouva"),
      viewWith(LOCALIZED),
    );
    const paragraph = retargeted.content.firstChild;
    expect(paragraph?.textContent).toBe("Smlouva");
    expect(paragraph?.attrs["outlineLevel"]).toBe(0);
  });

  test("leaves the id alone when the document defines no style at that level", () => {
    // The document has a level-1 heading but no level 3; inventing one would
    // mean writing a style definition as a side effect of a paste.
    expect(
      styleIdsOf(retargetPastedHeadingStyles(pastedHeading(3, "Detail"), viewWith(LOCALIZED))),
    ).toEqual(["Heading3"]);
  });

  test("leaves a style the document already defines alone", () => {
    const english: StyleDefinitions = {
      styles: [
        { styleId: "Heading1", type: "paragraph", name: "heading 1", pPr: { outlineLevel: 0 } },
        { styleId: "Nadpis1", type: "paragraph", name: "heading 1", pPr: { outlineLevel: 0 } },
      ],
    };
    expect(
      styleIdsOf(retargetPastedHeadingStyles(pastedHeading(1, "Title"), viewWith(english))),
    ).toEqual(["Heading1"]);
  });

  test("is a no-op without the document-styles plugin", () => {
    const slice = pastedHeading(1, "Title");
    expect(retargetPastedHeadingStyles(slice, viewWith(null))).toBe(slice);
  });

  test("leaves a non-heading paragraph alone", () => {
    const slice = new Slice(
      schema.node("doc", null, [
        schema.node("paragraph", { styleId: "SomeMissingStyle" }, [schema.text("Body")]),
      ]).content,
      0,
      0,
    );
    expect(retargetPastedHeadingStyles(slice, viewWith(LOCALIZED))).toBe(slice);
  });
});
