import { expect, test } from "bun:test";
import fixtures from "./fuzz-failure-class.fixtures.json";
import {
  classOfTitle,
  failureClass,
  normalizeFailureMessage,
  upgradeClassKey,
} from "./fuzz-failure-class";

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

test("a consumer-flow failure is one class across fixtures and modes", () => {
  const base = failureClass("consumer flow notes / direct", "unexpected item 1");
  expect(base.key).toBe('["consumer-flow","unexpected item <n>"]');
  expect(failureClass("consumer flow stories / direct", "unexpected item 2").key).toBe(base.key);
  expect(failureClass("consumer flow notes / suggested", "unexpected item 2").key).toBe(base.key);
  // The fixture and the mode stay on the class as row data.
  expect(failureClass("consumer flow stories / suggested", "unexpected item 2")).toMatchObject({
    fixture: "stories",
    mode: "suggested",
  });
  expect(failureClass("consumer flow notes / direct", "another failure").key).not.toBe(base.key);
  expect(failureClass("other notes / direct", "unexpected item 2").key).not.toBe(base.key);
});

test("other checks are classed and titled by the test that failed", () => {
  const message = "expect(received).toEqual(expected)";
  const first = failureClass("packages/a/one.property.test.ts::keeps structure", message);
  const second = failureClass("packages/a/one.property.test.ts::keeps numbering", message);
  const elsewhere = failureClass("packages/b/two.property.test.ts::keeps structure", message);
  expect(new Set([first.key, second.key, elsewhere.key]).size).toBe(3);
  expect(first.title).toBe(`Fuzz checks: keeps structure: ${message}`);
  expect(second.title).toBe(`Fuzz checks: keeps numbering: ${message}`);
  expect(failureClass("packages/a/one.property.test.ts::keeps structure", "other").key).not.toBe(
    first.key,
  );
  const long = failureClass(`file.test.ts::${"long name ".repeat(20)}`, message);
  expect(long.title.length).toBeLessThan(140);
  expect(long.title.endsWith(`…: ${message}`)).toBe(true);
  expect(failureClass("Table input", "text differs").title).toBe("Table input: text differs");
});

test("old and new titles of one consumer-flow failure are one class", () => {
  const message = "<text>: kind is <text>, expected <text>";
  const key = failureClass(
    "consumer flow lists / suggested",
    '"A": kind is "paragraph", expected "heading"',
  ).key;
  expect(
    classOfTitle(`Fuzz failure [73c7b924a911a13a]: consumer flow plain / direct: ${message}`)?.key,
  ).toBe(key);
  expect(classOfTitle(`Consumer flow: ${message}`)?.key).toBe(key);
  expect(classOfTitle("Consumer flow: something else")?.key).not.toBe(key);
  // A named family's own title is that family, whatever raised it.
  expect(classOfTitle("List numbering: listLevel mismatch")?.key).toBe(
    failureClass("consumer flow lists / direct", '"item": listLevel is 2, expected 3').key,
  );
  expect(classOfTitle("List numbering: something a person wrote")?.key).not.toBe(
    classOfTitle("Consumer flow: something a person wrote")?.key,
  );
});

test("class markers written with a fixture and a mode resolve to today's key", () => {
  expect(upgradeClassKey(["consumer-flow", "comments", "tracked-changes", "no comment"])).toBe(
    failureClass("consumer flow plain / direct", "step 2: no comment").key,
  );
  const name = "packages/a/one.property.test.ts::keeps structure";
  expect(upgradeClassKey(["<path>::keeps structure", name, "all", "boom"])).toBe(
    failureClass(name, "boom").key,
  );
  const current = ["consumer-flow", "no comment"];
  expect(upgradeClassKey(current)).toBe(JSON.stringify(current));
});
