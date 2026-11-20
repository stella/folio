import { expect, spyOn, test } from "bun:test";

import type { Document } from "../../model/document";
import { takeParagraphIds } from "../idLease";
import * as identities from "../ids";

test("allocation queries package membership and respects the lease and reserved spellings", () => {
  const document: Document = {
    package: {
      document: {
        content: [{ type: "paragraph", paraId: "0000000a", content: [] }],
      },
    },
  };
  const census = spyOn(identities, "packageParagraphIds");
  try {
    const options = {
      document,
      lease: { blocks: [{ kind: "paragraph", start: 10, end: 14 }] },
      reserved: new Set(["0000000c"]),
    } as const;
    expect(takeParagraphIds({ ...options, count: 2 })).toEqual([
      "0000000B",
      "0000000D",
    ]);
    expect(takeParagraphIds({ ...options, count: 3 })).toBeUndefined();
    expect(census).toHaveBeenCalledTimes(0);
  } finally {
    census.mockRestore();
  }
});
