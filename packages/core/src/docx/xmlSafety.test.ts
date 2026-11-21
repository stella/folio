import { describe, expect, spyOn, test } from "bun:test";

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

test("checking a longer text node adds no per-character JavaScript calls", () => {
  const count = (text: string) => {
    const units = spyOn(String.prototype, "charCodeAt");
    const points = spyOn(String.prototype, "codePointAt");
    try {
      const issue = getDocxXmlSafetyIssue(`<p>${text}</p>`);
      return {
        issue,
        units: units.mock.calls.length,
        points: points.mock.calls.length,
      };
    } finally {
      points.mockRestore();
      units.mockRestore();
    }
  };
  const text = "ř漢🌳";
  expect(count(text.repeat(1024))).toEqual(count(text));
  expect(getDocxXmlSafetyIssue(`<p>${text}\uD800</p>`)).toBe("not-well-formed");
  expect(getDocxXmlSafetyIssue(`<p>${text}\uFFFF</p>`)).toBe("not-well-formed");
  expect(getDocxXmlSafetyIssue(`<p>${text}</p>`)).toBeNull();
});
