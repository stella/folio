import { expect, test } from "bun:test";
import { assertExactModel } from "./exactModel";

test("exact journal oracle detects optional fields, package maps and media bytes", () => {
  expect(() => assertExactModel({}, { formatting: undefined })).toThrow();
  expect(() => assertExactModel({ preserved: true }, { preserved: false })).toThrow();
  expect(() => assertExactModel(new Map([["image", 1]]), new Map([["image", 2]]))).toThrow();
  expect(() => assertExactModel(Uint8Array.of(1).buffer, Uint8Array.of(2).buffer)).toThrow();
  assertExactModel(
    { formatting: undefined, media: new Map([["image", Uint8Array.of(1, 2).buffer]]) },
    { formatting: undefined, media: new Map([["image", Uint8Array.of(1, 2).buffer]]) },
  );
});
