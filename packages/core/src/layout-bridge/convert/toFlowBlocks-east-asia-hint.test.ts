import { describe, expect, test } from "bun:test";

import type { TextRun } from "../../layout-engine/types";
import { toProseDoc } from "../../prosemirror/conversion/toProseDoc";
import type { Document, StyleDefinitions, TextFormatting } from "../../types/document";
import { toFlowBlocks } from "./toFlowBlocks";

function firstTextRun(formatting: TextFormatting | undefined, styles?: StyleDefinitions): TextRun {
  const document: Document = {
    package: {
      ...(styles ? { styles } : {}),
      document: {
        content: [
          {
            type: "paragraph",
            content: [
              {
                type: "run",
                ...(formatting ? { formatting } : {}),
                content: [{ type: "text", text: "“A”" }],
              },
            ],
          },
        ],
      },
    },
  };
  const block = toFlowBlocks(toProseDoc(document, styles ? { styles } : {}), {})[0];
  if (block?.kind !== "paragraph") {
    throw new Error("expected a paragraph block");
  }
  const run = block.runs.find((candidate) => candidate.kind === "text");
  if (run?.kind !== "text") {
    throw new Error("expected a text run");
  }
  return run;
}

describe('w:rFonts w:hint="eastAsia" reaches layout', () => {
  test("a run's own hint", () => {
    const run = firstTextRun({ fontFamily: { eastAsia: "SimSun", hint: "eastAsia" } });
    expect(run.eastAsiaFontFamily).toBe("SimSun");
    expect(run.eastAsiaHint).toBe(true);
  });

  test("a run without a hint", () => {
    expect(firstTextRun({ fontFamily: { eastAsia: "SimSun" } }).eastAsiaHint).toBeUndefined();
  });

  test("a hint from the document defaults, and a run's w:hint=default over it", () => {
    const styles: StyleDefinitions = {
      docDefaults: { rPr: { fontFamily: { eastAsia: "SimSun", hint: "eastAsia" } } },
      styles: [],
    };
    expect(firstTextRun(undefined, styles).eastAsiaHint).toBe(true);
    expect(firstTextRun({ fontFamily: { hint: "default" } }, styles).eastAsiaHint).toBe(false);
  });
});
