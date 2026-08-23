import { describe, expect, test } from "bun:test";

import { paragraphToStyle } from "./formatToStyle";

describe("paragraphToStyle indentation", () => {
  test("hangingIndent wins when both first-line fields are present", () => {
    expect(paragraphToStyle({ indentFirstLine: 720, hangingIndent: true }).textIndent).toBe(
      "-48px",
    );
  });

  test("keeps a regular first-line indent positive", () => {
    expect(paragraphToStyle({ indentFirstLine: 720 }).textIndent).toBe("48px");
  });
});
