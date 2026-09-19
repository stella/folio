import { describe, expect, test } from "bun:test";
import { MAX_RETAINED_PARSE_WARNINGS_PER_CODE, PARSE_WARNING_CODES } from "@stll/docx-core/model";

import { createParseWarningCollector } from "./parseContext";

describe("parse warning collector", () => {
  test("a scoped context reports against its own part", () => {
    const { context, warnings } = createParseWarningCollector("word/document.xml");

    context.scoped({ part: "word/footnotes.xml", at: "w:id 3" }).warn({
      code: PARSE_WARNING_CODES.duplicateNoteId,
    });

    expect(warnings()).toEqual([
      {
        code: PARSE_WARNING_CODES.duplicateNoteId,
        location: { part: "word/footnotes.xml", at: "w:id 3" },
        count: 1,
      },
    ]);
  });

  test("a pathological file cannot balloon the list: past the cap it is counted", () => {
    const { context, warnings } = createParseWarningCollector("word/styles.xml");
    const emitted = MAX_RETAINED_PARSE_WARNINGS_PER_CODE + 5;

    for (let index = 0; index < emitted; index += 1) {
      context.warn({
        code: PARSE_WARNING_CODES.unrecognisedOnOffValue,
        value: `v${String(index)}`,
      });
    }

    const recorded = warnings();
    expect(recorded).toHaveLength(MAX_RETAINED_PARSE_WARNINGS_PER_CODE + 1);
    const overflow = recorded.at(-1);
    expect(overflow?.count).toBe(5);
    expect(overflow?.detail).toContain("not retained");
    // Nothing is lost: retained plus counted equals what happened.
    expect(recorded.reduce((total, warning) => total + warning.count, 0)).toBe(emitted);
  });

  test("one code overflowing does not spend another code's budget", () => {
    const { context, warnings } = createParseWarningCollector();

    for (let index = 0; index <= MAX_RETAINED_PARSE_WARNINGS_PER_CODE; index += 1) {
      context.warn({ code: PARSE_WARNING_CODES.unrecognisedOnOffValue });
    }
    context.warn({ code: PARSE_WARNING_CODES.borderWithoutValue });

    const byCode = warnings().filter(
      (warning) => warning.code === PARSE_WARNING_CODES.borderWithoutValue,
    );
    expect(byCode).toHaveLength(1);
  });
});
