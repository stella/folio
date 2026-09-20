import { describe, expect, test } from "bun:test";

import { TABLE_PROPERTY_JUSTIFICATIONS, toTablePropertyJustification } from "./tableOperations";

describe("editable table placement", () => {
  test("keeps every schema placement available to property dialogs", () => {
    expect(TABLE_PROPERTY_JUSTIFICATIONS).toEqual(["left", "center", "right", "start", "end"]);
    expect(TABLE_PROPERTY_JUSTIFICATIONS.map(toTablePropertyJustification)).toEqual(
      TABLE_PROPERTY_JUSTIFICATIONS,
    );
  });

  test.each([undefined, null, "both"])('falls back from "$p" to left', (value) => {
    expect(toTablePropertyJustification(value)).toBe("left");
  });
});
