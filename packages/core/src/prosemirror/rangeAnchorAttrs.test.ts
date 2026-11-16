import { describe, expect, test } from "bun:test";

import {
  readRangeAnchorAttrs,
  readRangeAnchorValues,
} from "./rangeAnchorAttrs";
import { schema } from "./schema";

describe("range anchor attrs", () => {
  test.each(
    [
      [],
      [{ type: "commentRangeStart", id: 1 }],
      [
        { type: "commentRangeEnd", id: 1 },
        { type: "commentRangeStart", id: 1 },
      ],
      [
        { type: "commentRangeStart", id: 1 },
        { type: "commentRangeEnd", id: 2 },
      ],
      [
        { type: "commentRangeStart", id: 1 },
        { type: "moveToRangeStart", id: 2, author: "Reviewer", name: "Move" },
        { type: "commentRangeEnd", id: 1 },
        { type: "moveToRangeEnd", id: 2 },
      ],
    ].map((markers) => ({ markers })),
  )(
    "rejects an empty, unbalanced or unsupported mixed cluster",
    ({ markers }) => {
      expect(readRangeAnchorValues(markers).ok).toBe(false);
    },
  );

  test("rejects malformed fields before the DOCX serializer can consume them", () => {
    const node = schema.node("rangeAnchor", {
      markers: [
        {
          type: "moveFromRangeStart",
          id: -1,
          name: "",
          author: "",
          date: null,
          displacedByCustomXml: "middle",
          colFirst: -2,
          colLast: null,
        },
        {
          type: "moveFromRangeEnd",
          id: -1,
          displacedByCustomXml: "middle",
        },
      ],
    });

    const result = readRangeAnchorAttrs(node);

    expect(result.ok).toBe(false);
    if (result.ok) {
      throw new Error("Expected malformed range anchor attrs to be rejected");
    }
    expect(result.issues.map(({ path }) => path)).toEqual(
      expect.arrayContaining([
        "rangeAnchor.attrs.markers[0].id",
        "rangeAnchor.attrs.markers[0].name",
        "rangeAnchor.attrs.markers[0].author",
        "rangeAnchor.attrs.markers[0].date",
        "rangeAnchor.attrs.markers[0].displacedByCustomXml",
        "rangeAnchor.attrs.markers[0].colFirst",
        "rangeAnchor.attrs.markers[0].colLast",
        "rangeAnchor.attrs.markers[1].displacedByCustomXml",
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
    const node = schema.node("rangeAnchor", { markers: [start, end] });

    const result = readRangeAnchorAttrs(node);

    expect(result).toEqual({ ok: true, value: { markers: [start, end] } });
  });
});
