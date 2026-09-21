import { describe, expect, test } from "bun:test";

import { parseQuery } from "./container-contract-doc-counts";

describe("container contract count filters", () => {
  test("accepts declared fields and exact comparison grammar", () => {
    expect(parseQuery("kind=child&reason!=neverParsed,replayOnly")).toEqual([
      { field: "kind", negated: false, values: ["child"] },
      { field: "reason", negated: true, values: ["neverParsed", "replayOnly"] },
    ]);
  });

  test.each(["toString=child", "kind=child=stale", "kind!=child!=stale", "kind="])(
    "rejects an inherited field or malformed comparison: %s",
    (query) => {
      expect(() => parseQuery(query)).toThrow();
    },
  );
});
