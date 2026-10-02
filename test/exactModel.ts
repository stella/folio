import { deepStrictEqual } from "node:assert/strict";

/** Journal equality includes own undefined fields, Maps, media bytes and identities. */
export const assertExactModel = (actual: unknown, expected: unknown): void => {
  deepStrictEqual(actual, expected);
};
