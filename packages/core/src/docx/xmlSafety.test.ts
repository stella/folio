import { describe, expect, test } from "bun:test";

import { getDocxXmlSafetyIssue } from "./xmlSafety";

describe("getDocxXmlSafetyIssue", () => {
  test.each([
    '<?xml version = "1.0"?><root/>',
    '\uFEFF<?xml version = "1.0" encoding = \'UTF-8\' standalone = "yes"?><root/>',
  ])("accepts a legal XML declaration: %s", (xml) => {
    expect(getDocxXmlSafetyIssue(xml)).toBeNull();
  });

  test.each([
    ' \uFEFF<?xml version="1.0"?><root/>',
    "<root>\uFEFF</root>",
    '<!--before--><?xml version="1.0"?><root/>',
  ])("rejects a misplaced declaration marker: %s", (xml) => {
    expect(getDocxXmlSafetyIssue(xml)).toBe("not-well-formed");
  });
});
