import { describe, expect, test } from "bun:test";

import { schema } from "./schema";
import { paragraphRunStyleContext } from "./runStyleFormatting";

describe("paragraphRunStyleContext", () => {
  test("discards a table run base after the paragraph style changes", () => {
    const paragraph = schema.node("paragraph", {
      styleId: "Target",
      _tableRunFormatting: { fontSize: 24 },
    });
    const styleResolver = {
      getDefaultCharacterStyle: () => undefined,
      getDefaultParagraphStyle: () => undefined,
      getDocDefaults: () => undefined,
      getRunStyleOwnProperties: () => undefined,
      getStyle: () => undefined,
      resolveParagraphStyle: () => ({ runFormatting: { fontSize: 28 } }),
    };

    expect(paragraphRunStyleContext(paragraph, styleResolver).baseParagraphFormatting).toEqual({
      fontSize: 28,
    });
  });
});
