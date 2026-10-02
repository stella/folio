import { expect, test } from "bun:test";
import fixtures from "./fuzz-failure-class.fixtures.json";
import { classOfTitle, failureClass, normalizeFailureMessage } from "./fuzz-failure-class";

test("the eighteen existing titles have stable reporting classes", () => {
  expect(fixtures).toHaveLength(18);
  for (const { title, key } of fixtures) expect(classOfTitle(title)?.key).toBe(key);
});

test("volatile assertion data does not split a reporting class", () => {
  expect(
    normalizeFailureMessage('step 42: id 12ab34cd at /tmp/run-8/file.xml offset -32 says "one"'),
  ).toBe(
    normalizeFailureMessage('step 99: id 56ab78cd at /tmp/run-9/other.xml offset -64 says "two"'),
  );
  expect(
    failureClass("consumer flow lists / direct", '"item": listLevel is 2, expected 3').key,
  ).toBe(
    failureClass("consumer flow emoji / suggested", '"other": listLevel is undefined, expected 4')
      .key,
  );
});

test("unknown assertions retain fixture, mode, and failure kind boundaries", () => {
  const base = failureClass("consumer flow notes / direct", "unexpected item 1").key;
  expect(failureClass("consumer flow stories / direct", "unexpected item 2").key).not.toBe(base);
  expect(failureClass("consumer flow notes / suggested", "unexpected item 2").key).not.toBe(base);
  expect(failureClass("other notes / direct", "unexpected item 2").key).not.toBe(base);
});
