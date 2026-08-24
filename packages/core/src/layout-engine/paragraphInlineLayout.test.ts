import { describe, expect, test } from "bun:test";

import { resolvePhysicalParagraphInlineLayout } from "../utils/paragraphInlineLayout";
import type { ParagraphAttrs } from "./types";

const resolve = (attrs?: ParagraphAttrs, text = "text") =>
  resolvePhysicalParagraphInlineLayout({
    kind: "paragraph",
    id: "paragraph",
    runs: [{ kind: "text", text, rtl: text === "مرحبا" }],
    attrs,
  });

describe("physical paragraph inline layout", () => {
  test("mirrors explicit horizontal alignment and asymmetric indents under bidi", () => {
    expect(
      resolve({ alignment: "right", bidi: true, indent: { left: 12, right: 34 } }),
    ).toMatchObject({
      alignment: "left",
      explicitAlignment: "left",
      indentLeft: 34,
      indentRight: 12,
      isRtl: true,
    });
    expect(resolve({ alignment: "left", bidi: true })).toMatchObject({
      alignment: "right",
      explicitAlignment: "right",
    });
  });

  test("preserves center and justify under bidi", () => {
    expect(resolve({ alignment: "center", bidi: true }).alignment).toBe("center");
    expect(resolve({ alignment: "justify", bidi: true }).alignment).toBe("justify");
  });

  test("does not mirror explicit false or inferred RTL paragraphs", () => {
    expect(
      resolve({ alignment: "right", bidi: false, indent: { left: 12, right: 34 } }),
    ).toMatchObject({ alignment: "right", indentLeft: 12, indentRight: 34, isRtl: false });
    expect(resolve({ alignment: "left", indent: { left: 12, right: 34 } }, "مرحبا")).toMatchObject({
      alignment: "left",
      indentLeft: 12,
      indentRight: 34,
      isRtl: true,
    });
  });

  test("defaults RTL paragraphs to physical right without inventing explicit alignment", () => {
    expect(resolve({ bidi: true })).toEqual({
      alignment: "right",
      indentLeft: 0,
      indentRight: 0,
      isRtl: true,
    });
  });
});
