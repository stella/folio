import { describe, expect, test } from "bun:test";

import { schema } from "./schema";
import { readRangeAnchorAttrs } from "./rangeAnchorAttrs";

describe("range anchor attrs", () => {
  test("rejects malformed fields before the DOCX serializer can consume them", () => {
    const node = schema.node("rangeAnchor", {
      start: {
        type: "moveFromRangeStart",
        id: -1,
        name: "",
        author: "",
        date: null,
        displacedByCustomXml: "middle",
        colFirst: -2,
        colLast: null,
      },
      end: {
        type: "moveFromRangeEnd",
        id: -1,
        displacedByCustomXml: "middle",
      },
    });

    const result = readRangeAnchorAttrs(node);

    expect(result.ok).toBe(false);
    if (result.ok) {
      throw new Error("Expected malformed range anchor attrs to be rejected");
    }
    expect(result.issues.map(({ path }) => path)).toEqual(
      expect.arrayContaining([
        "rangeAnchor.attrs.start.id",
        "rangeAnchor.attrs.start.name",
        "rangeAnchor.attrs.start.author",
        "rangeAnchor.attrs.start.date",
        "rangeAnchor.attrs.start.displacedByCustomXml",
        "rangeAnchor.attrs.start.colFirst",
        "rangeAnchor.attrs.start.colLast",
        "rangeAnchor.attrs.end.displacedByCustomXml",
      ]),
    );
  });

  test("accepts and preserves every serialized move-range field", () => {
    const start = {
      type: "moveToRangeStart" as const,
      id: 4,
      name: "destination",
      author: "Reviewer",
      date: "2026-01-01T00:00:00Z",
      displacedByCustomXml: "next" as const,
      colFirst: 2,
      colLast: 5,
    };
    const end = {
      type: "moveToRangeEnd" as const,
      id: 4,
      displacedByCustomXml: "prev" as const,
    };
    const node = schema.node("rangeAnchor", { start, end });

    const result = readRangeAnchorAttrs(node);

    expect(result).toEqual({ ok: true, value: { start, end } });
  });
});
