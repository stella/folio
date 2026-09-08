import { expect, test } from "bun:test";

import type { ParagraphPropertyChange } from "../../types/document";
import { serializeParagraphFormatting } from "./paragraphSerializer";

const propertyChange = (id: number): ParagraphPropertyChange => ({
  type: "paragraphPropertyChange",
  info: { id, author: "Reviewer", date: "2026-09-08T00:00:00.000Z" },
  previousFormatting: { spaceBefore: id * 120 },
});

test("refuses to serialize multiple non-suggested sibling w:pPrChange elements", () => {
  expect(() =>
    serializeParagraphFormatting({ spaceBefore: 360 }, [propertyChange(1), propertyChange(2)]),
  ).toThrow("A paragraph cannot serialize more than one w:pPrChange");
});
