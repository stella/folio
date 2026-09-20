import { describe, expect, test } from "bun:test";

import type { SelectionState } from "@stll/folio-core/prosemirror";
import { NO_LIST_STATE } from "@stll/folio-core/prosemirror";

import { buildSelectionFormatting } from "./selectionFormattingBuilder";

function makeSelection(
  textFormatting: Partial<SelectionState["textFormatting"]> = {},
  paragraphFormatting: Partial<SelectionState["paragraphFormatting"]> = {},
  styleId: string | null = null,
): SelectionState {
  return {
    hasSelection: false,
    isMultiParagraph: false,
    textFormatting,
    paragraphFormatting,
    styleId,
    startParagraphIndex: 0,
    endParagraphIndex: 0,
    listState: NO_LIST_STATE,
  };
}

describe("buildSelectionFormatting", () => {
  test("emits the always-present derived booleans", () => {
    const formatting = buildSelectionFormatting({
      selectionState: makeSelection(),
      fontFamily: undefined,
      fontSize: undefined,
      textColor: undefined,
      listState: undefined,
    });
    expect(formatting).toEqual({
      underline: false,
      superscript: false,
      subscript: false,
      bidi: false,
    });
  });

  test("derives superscript / subscript from vertAlign", () => {
    expect(
      buildSelectionFormatting({
        selectionState: makeSelection({ vertAlign: "superscript" }),
        fontFamily: undefined,
        fontSize: undefined,
        textColor: undefined,
        listState: undefined,
      }).superscript,
    ).toBe(true);
    expect(
      buildSelectionFormatting({
        selectionState: makeSelection({ vertAlign: "subscript" }),
        fontFamily: undefined,
        fontSize: undefined,
        textColor: undefined,
        listState: undefined,
      }).subscript,
    ).toBe(true);
  });

  test("copies optional fields only when their source is defined", () => {
    const formatting = buildSelectionFormatting({
      selectionState: makeSelection({ bold: true, italic: false }, { alignment: "center" }),
      fontFamily: undefined,
      fontSize: undefined,
      textColor: undefined,
      listState: undefined,
    });
    expect(formatting.bold).toBe(true);
    expect(formatting.italic).toBe(false);
    expect(formatting.alignment).toBe("center");
    // Fields whose sources were `undefined` must be absent (not `undefined`)
    // so the prop shape stays compatible with exactOptionalPropertyTypes.
    expect("strike" in formatting).toBe(false);
    expect("fontFamily" in formatting).toBe(false);
    expect("fontSize" in formatting).toBe(false);
  });

  test("passes through resolved font / color / list state", () => {
    const formatting = buildSelectionFormatting({
      selectionState: makeSelection(),
      fontFamily: "Arimo",
      fontSize: 22,
      textColor: "var(--test-color)",
      listState: { type: "bullet", level: 0, isInList: true, numId: 1 },
    });
    expect(formatting.fontFamily).toBe("Arimo");
    expect(formatting.fontSize).toBe(22);
    expect(formatting.color).toBe("var(--test-color)");
    expect(formatting.listState).toEqual({
      type: "bullet",
      level: 0,
      isInList: true,
      numId: 1,
    });
  });

  test("copies the paragraph styleId for any non-null source value", () => {
    expect(
      buildSelectionFormatting({
        selectionState: makeSelection({}, {}, "Heading1"),
        fontFamily: undefined,
        fontSize: undefined,
        textColor: undefined,
        listState: undefined,
      }).styleId,
    ).toBe("Heading1");
    // Empty string is technically permitted by `SelectionState.styleId`
    // (`string | null`) and must be passed through — only `null` means
    // "no paragraph style".
    expect(
      buildSelectionFormatting({
        selectionState: makeSelection({}, {}, ""),
        fontFamily: undefined,
        fontSize: undefined,
        textColor: undefined,
        listState: undefined,
      }).styleId,
    ).toBe("");
    // `null` styleId must not appear on the result.
    expect(
      "styleId" in
        buildSelectionFormatting({
          selectionState: makeSelection({}, {}, null),
          fontFamily: undefined,
          fontSize: undefined,
          textColor: undefined,
          listState: undefined,
        }),
    ).toBe(false);
  });
});
