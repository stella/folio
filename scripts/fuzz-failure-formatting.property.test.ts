import { expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import { assertProperty, propertyTestTimeout } from "../test/property-testing";
import { failureMarker, diffShape } from "../test/consumer-scenarios/support/failure-fingerprints";
import { failureClass } from "./fuzz-failure-class";

setDefaultTimeout(propertyTestTimeout(5_000));

const marker = (failure: Error) =>
  failureMarker({
    test: "generated comparison",
    seed: 11,
    repro: "bun replay.ts",
    failure,
    flow: "typing",
  });

test("failure identity and diff shape are invariant under ANSI decoration", () => {
  assertProperty(
    fc.property(
      fc.constantFrom(
        'expect(received).toEqual(expected)\n[2].text: "before" → "after"',
        "step 42: bold changed at 12, outside what was asked",
        'step 3: the result is not what was asked\n  no comment {"id":1234}',
      ),
      fc.constantFrom("\u001b[31m", "\u001b[1;32m", "\u001b]8;;https://example.invalid/\u0007"),
      (message, decoration) => {
        const close = decoration.startsWith("\u001b]") ? "\u001b]8;;\u0007" : "\u001b[0m";
        const decorated = Array.from(message, (character) => decoration + character + close).join(
          "",
        );
        const plain = marker(new Error(message));
        const colored = marker(new Error(decorated));
        expect(colored).toEqual(plain);
        expect(diffShape(new Error(decorated))).toBe(diffShape(new Error(message)));
        expect(failureClass(plain.test, decorated).key).toBe(failureClass(plain.test, message).key);
      },
    ),
    { numRuns: 50 },
  );
});
