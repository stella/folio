import { describe, expect, test } from "bun:test";

import type { DocumentBody } from "../types/document";
import { normalizeRenderedPageBreakHints } from "./renderedPageBreakNormalization";

describe("rendered page-break normalization", () => {
  test("keeps only hints whose leading marker still precedes serializable content", () => {
    const documentBody = {
      content: [
        {
          type: "paragraph",
          renderedPageBreakBefore: true,
          content: [{ type: "run", content: [{ type: "renderedPageBreak" }] }],
        },
        {
          type: "paragraph",
          renderedPageBreakBefore: true,
          content: [
            {
              type: "run",
              content: [{ type: "renderedPageBreak" }, { type: "text", text: "retained" }],
            },
          ],
        },
        {
          type: "paragraph",
          renderedPageBreakBefore: true,
          content: [
            {
              type: "run",
              content: [{ type: "text", text: "before" }, { type: "renderedPageBreak" }],
            },
          ],
        },
      ],
    } satisfies DocumentBody;

    normalizeRenderedPageBreakHints({ documentBody });

    expect(
      documentBody.content.map(({ renderedPageBreakBefore }) => renderedPageBreakBefore),
    ).toEqual([undefined, true, undefined]);
  });
});
