import { expect, spyOn, test } from "bun:test";

import type { Insertion } from "../../model/document";
import { identityKeysIn } from "../ids";

test("the identity census reads a record list without mapping it into intermediates", () => {
  const size = 256;
  const inserts = Array.from(
    { length: size },
    (_, index) =>
      ({
        type: "insertion",
        info: { id: index + 1, author: "Editor" },
        content: [{ type: "run", content: [{ type: "text", text: "Text" }] }],
      }) satisfies Insertion,
  );
  const intermediate = spyOn(inserts, "map");
  try {
    expect(identityKeysIn(inserts)).toEqual(
      Array.from({ length: size }, (_, index) => `revision:${index + 1}`),
    );
    expect(intermediate).toHaveBeenCalledTimes(0);
  } finally {
    intermediate.mockRestore();
  }
});
