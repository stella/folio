import { describe, expect, test } from "bun:test";

import type { PositionedBookmarkMarker } from "../types/document";

import { schema } from "./schema";
import { assertValidProseMirrorDocument, validateProseMirrorDocument } from "./validation";

/**
 * Assert an issue whose message starts with `fragment`. Bookmark messages carry
 * the bookmark's name and its paragraph so a failing document is diagnosable;
 * the tests pin the reason, not that decoration.
 */
const expectIssueContaining = (
  result: ReturnType<typeof validateProseMirrorDocument>,
  fragment: string,
): void => {
  expect(result.issues.map((issue) => issue.message)).toEqual(
    expect.arrayContaining([expect.stringContaining(fragment)]),
  );
};

describe("ProseMirror document validation", () => {
  test("reports attr issues with document paths", () => {
    const highlight = schema.mark("highlight", { color: "customYellow" });
    const doc = schema.node("doc", null, [
      schema.node("paragraph", { paraId: 12 }, [
        schema.text("bad", [highlight]),
        schema.node("field", {
          fieldType: "NOT_A_FIELD",
          instruction: " PAGE ",
          displayText: "1",
          fieldKind: "simple",
        }),
      ]),
    ]);

    const result = validateProseMirrorDocument(doc);

    expect(result.valid).toBe(false);
    expect(result.issues.map((issue) => issue.path)).toEqual(
      expect.arrayContaining([
        "doc.content[0].paragraph.attrs.paraId",
        "doc.content[0].content[0].marks[0].highlight.attrs.color",
        "doc.content[0].content[1].field.attrs.fieldType",
      ]),
    );
  });

  test("throws a formatted validation error", () => {
    const doc = schema.node("doc", null, [
      schema.node("paragraph", { lineSpacing: "240" }, [schema.text("invalid")]),
    ]);

    expect(() => assertValidProseMirrorDocument(doc, "Cannot use invalid PM document")).toThrow(
      "ProseMirror document error at doc.content[0].paragraph.attrs.lineSpacing",
    );
  });

  test.each([
    {
      name: "lone start",
      content: [schema.node("bookmarkBoundary", { type: "start", id: 1, name: "one" })],
      message: 'Bookmark "one" (id 1) has no matching end boundary',
    },
    {
      name: "lone end",
      content: [schema.node("bookmarkBoundary", { type: "end", id: 1 })],
      message: "Bookmark id 1 has no open start boundary",
    },
    {
      name: "duplicate starts",
      content: [
        schema.node("bookmarkBoundary", { type: "start", id: 1, name: "one" }),
        schema.node("bookmarkBoundary", { type: "end", id: 1 }),
        schema.node("bookmarkBoundary", { type: "start", id: 1, name: "duplicate" }),
        schema.node("bookmarkBoundary", { type: "end", id: 1 }),
      ],
      message: 'Bookmark "duplicate" (id 1) has more than one start boundary',
    },
  ])("rejects $name bookmark boundaries", ({ content, message }) => {
    const doc = schema.node("doc", null, [schema.node("paragraph", null, content)]);

    const result = validateProseMirrorDocument(doc);

    expect(result.valid).toBe(false);
    expectIssueContaining(result, message);
  });

  test("tolerates one id repeated across paragraph bookmark attrs", () => {
    // Real documents repeat a bookmark id this way (a stray `_GoBack`, a
    // template assembled from several sources). Word and LibreOffice keep the
    // first start and open the file; refusing it here made those documents
    // unloadable. The attr serializes as a start and an end around its own
    // paragraph, so the repeat writes back exactly what was read.
    const paragraph = (paraId: string) =>
      schema.node("paragraph", { paraId, bookmarks: [{ id: 0, name: "_GoBack" }] }, [
        schema.text(`clause ${paraId}`),
      ]);
    const doc = schema.node("doc", null, [paragraph("A0000001"), paragraph("A0000002")]);

    expect(validateProseMirrorDocument(doc)).toEqual({ valid: true, issues: [] });
  });

  test("names the bookmark and the paragraph a boundary failure sits in", () => {
    const doc = schema.node("doc", null, [
      schema.node("paragraph", { paraId: "B0000001" }, [
        schema.node("bookmarkBoundary", { type: "start", id: 4, name: "unclosed" }),
        schema.text("body"),
      ]),
    ]);

    expectIssueContaining(
      validateProseMirrorDocument(doc),
      'Bookmark "unclosed" (id 4) has no matching end boundary (paragraph B0000001)',
    );
  });

  test("allows paired bookmark ranges to overlap", () => {
    const doc = schema.node("doc", null, [
      schema.node("paragraph", null, [
        schema.node("bookmarkBoundary", { type: "start", id: 1, name: "one" }),
        schema.node("bookmarkBoundary", { type: "start", id: 2, name: "two" }),
        schema.node("bookmarkBoundary", { type: "end", id: 1 }),
        schema.node("bookmarkBoundary", { type: "end", id: 2 }),
      ]),
    ]);

    expect(validateProseMirrorDocument(doc)).toEqual({ valid: true, issues: [] });
  });

  test("pairs table and row bookmark starts with ends inside following cells", () => {
    const paragraph = schema.node("paragraph", null, [
      schema.node("bookmarkBoundary", { type: "end", id: 1 }),
      schema.node("bookmarkBoundary", { type: "end", id: 2 }),
    ]);
    const row = schema.node(
      "tableRow",
      {
        _bookmarks: [{ index: 0, marker: { type: "bookmarkStart", id: 2, name: "row range" } }],
      },
      [schema.node("tableCell", null, [paragraph])],
    );
    const table = schema.node(
      "table",
      {
        _bookmarks: [{ index: 0, marker: { type: "bookmarkStart", id: 1, name: "table range" } }],
      },
      [row],
    );
    const doc = schema.node("doc", null, [table]);

    expect(validateProseMirrorDocument(doc)).toEqual({ valid: true, issues: [] });
  });

  test("allows an unmatched positioned bookmark end from an imported table", () => {
    const row = schema.node(
      "tableRow",
      {
        _bookmarks: [{ index: 0, marker: { type: "bookmarkEnd", id: 1 } }],
      },
      [schema.node("tableCell", null, [schema.node("paragraph")])],
    );
    const doc = schema.node("doc", null, [schema.node("table", null, [row])]);

    expect(validateProseMirrorDocument(doc)).toEqual({ valid: true, issues: [] });
  });

  const tableWithRowBookmark = (index: number, id: number) =>
    schema.node("doc", null, [
      schema.node("table", null, [
        schema.node(
          "tableRow",
          {
            _bookmarks: [{ index, marker: { type: "bookmarkStart", id, name: "range" } }],
          },
          [schema.node("tableCell", null, [schema.node("paragraph")])],
        ),
      ]),
    ]);

  test.each([
    { label: "fractional", value: 0.5 },
    { label: "negative", value: -1 },
    { label: "unsafe", value: Number.MAX_SAFE_INTEGER + 1 },
  ])("rejects a $label positioned bookmark index", ({ value }) => {
    const result = validateProseMirrorDocument(tableWithRowBookmark(value, 1));

    expect(result.valid).toBe(false);
    expect(result.issues).toContainEqual({
      path: "doc.content[0].content[0].tableRow.attrs._bookmarks[0].index",
      message: "Expected a non-negative safe integer.",
    });
  });

  test.each([
    { label: "fractional", value: 0.5 },
    { label: "negative", value: -1 },
    { label: "unsafe", value: Number.MAX_SAFE_INTEGER + 1 },
  ])("rejects a $label positioned bookmark id", ({ value }) => {
    const result = validateProseMirrorDocument(tableWithRowBookmark(0, value));

    expect(result.valid).toBe(false);
    expect(result.issues).toContainEqual({
      path: "doc.content[0].content[0].tableRow.attrs._bookmarks[0].marker.id",
      message: "Expected a non-negative safe integer.",
    });
  });

  test("rejects a positioned bookmark beyond its owning node children", () => {
    const result = validateProseMirrorDocument(tableWithRowBookmark(2, 1));

    expect(result.valid).toBe(false);
    expect(result.issues).toContainEqual({
      path: "doc.content[0].content[0].tableRow.attrs._bookmarks[0].index",
      message: "Expected a child position between 0 and 1.",
    });
  });

  test("scans positioned bookmarks a constant number of times for a large table row", () => {
    const childCount = 256;
    let entriesCalls = 0;
    const bookmarks = new Proxy(
      Array.from(
        { length: childCount },
        (_, index) =>
          [
            { index, marker: { type: "bookmarkStart", id: index, name: `range-${index}` } },
            { index, marker: { type: "bookmarkEnd", id: index } },
          ] satisfies PositionedBookmarkMarker[],
      ).flat(),
      {
        get: (target, property, receiver) => {
          if (property === "entries") {
            entriesCalls += 1;
          }
          return Reflect.get(target, property, receiver);
        },
      },
    );
    const cells = Array.from({ length: childCount }, () =>
      schema.node("tableCell", null, [schema.node("paragraph")]),
    );
    const row = schema.node("tableRow", { _bookmarks: bookmarks }, cells);
    const doc = schema.node("doc", null, [schema.node("table", null, [row])]);

    expect(validateProseMirrorDocument(doc)).toEqual({ valid: true, issues: [] });
    expect(entriesCalls).toBeLessThanOrEqual(3);
  });

  test("allows a bookmark to overlap a tracked hyperlink", () => {
    const hyperlink = schema.mark("hyperlink", {
      href: "https://example.test",
      _docxHyperlinkIndex: 1,
    });
    const insertion = schema.mark("insertion", {
      revisionId: 7,
      author: "Reviewer",
      date: null,
      initials: null,
      moveKind: null,
    });
    const marks = [hyperlink, insertion];
    const doc = schema.node("doc", null, [
      schema.node("paragraph", null, [
        schema.node("bookmarkBoundary", { type: "start", id: 1, name: "one" }, null, marks),
        schema.text("outside"),
        schema.node("bookmarkBoundary", { type: "end", id: 1 }, null, marks),
      ]),
    ]);

    const result = validateProseMirrorDocument(doc);

    expect(result).toEqual({ valid: true, issues: [] });
  });

  test("allows a bookmark boundary directly inside a tracked change", () => {
    const insertion = schema.mark("insertion", {
      revisionId: 7,
      author: "Reviewer",
      date: null,
      initials: null,
      moveKind: null,
    });
    const doc = schema.node("doc", null, [
      schema.node("paragraph", null, [
        schema.node("bookmarkBoundary", { type: "start", id: 1, name: "one" }, null, [insertion]),
        schema.text("tracked", [insertion]),
        schema.node("bookmarkBoundary", { type: "end", id: 1 }, null, [insertion]),
      ]),
    ]);

    const result = validateProseMirrorDocument(doc);

    expect(result).toEqual({ valid: true, issues: [] });
  });

  test("rejects a field boundary without its required hyperlink parent", () => {
    const doc = schema.node("doc", null, [
      schema.node("paragraph", null, [
        schema.node(
          "structuredField",
          {
            fieldType: "REF",
            instruction: " REF target ",
            displayText: "Target",
            fieldKind: "simple",
          },
          [
            schema.node("bookmarkBoundary", { type: "start", id: 3, name: "target" }),
            schema.text("Target"),
            schema.node("bookmarkBoundary", { type: "end", id: 3 }),
          ],
        ),
      ]),
    ]);

    expect(validateProseMirrorDocument(doc).issues.map((issue) => issue.message)).toContain(
      "Bookmark boundaries inside fields require a hyperlink parent.",
    );
  });

  test("rejects children forged onto an ordinary field leaf", () => {
    const field = schema.nodes.field.create(
      {
        fieldType: "PAGE",
        instruction: " PAGE ",
        displayText: "1",
        fieldKind: "complex",
      },
      schema.text("forged"),
    );
    const doc = schema.node("doc", null, [schema.node("paragraph", null, [field])]);

    expect(validateProseMirrorDocument(doc).issues.map((issue) => issue.message)).toContain(
      "Ordinary fields cannot contain structured result children.",
    );
  });

  test("rejects any structured children in complex field results", () => {
    const doc = schema.node("doc", null, [
      schema.node("paragraph", null, [
        schema.node(
          "structuredField",
          {
            fieldType: "PAGE",
            instruction: " PAGE ",
            displayText: "1",
            fieldKind: "complex",
          },
          [schema.text("1")],
        ),
      ]),
    ]);

    expect(validateProseMirrorDocument(doc).issues.map((issue) => issue.message)).toContain(
      "Complex fields cannot contain structured result children.",
    );
  });

  test("rejects structured simple field children without a hyperlink", () => {
    const doc = schema.node("doc", null, [
      schema.node("paragraph", null, [
        schema.node(
          "structuredField",
          {
            fieldType: "REF",
            instruction: " REF target ",
            displayText: "Target",
            fieldKind: "simple",
          },
          [schema.text("Target")],
        ),
      ]),
    ]);

    expect(validateProseMirrorDocument(doc).issues.map((issue) => issue.message)).toContain(
      "Structured simple fields require hyperlink, page-break, preserved or wrapper content.",
    );
  });

  test("rejects ids shared by legacy paragraph bookmarks and boundary atoms", () => {
    const doc = schema.node("doc", null, [
      schema.node("paragraph", { bookmarks: [{ id: 1, name: "legacy" }] }, [
        schema.node("bookmarkBoundary", { type: "start", id: 1, name: "pasted" }),
        schema.text("collision"),
        schema.node("bookmarkBoundary", { type: "end", id: 1 }),
      ]),
    ]);

    const result = validateProseMirrorDocument(doc);

    expect(result.valid).toBe(false);
    expectIssueContaining(result, 'Bookmark "pasted" (id 1) has more than one start boundary');
  });

  test("rejects an internally pasted pair that collides with an existing id", () => {
    const pair = (name: string) =>
      schema.node("paragraph", null, [
        schema.node("bookmarkBoundary", { type: "start", id: 1, name }),
        schema.text(name),
        schema.node("bookmarkBoundary", { type: "end", id: 1 }),
      ]);
    const doc = schema.node("doc", null, [pair("existing"), pair("pasted")]);

    const result = validateProseMirrorDocument(doc);

    expect(result.valid).toBe(false);
    expectIssueContaining(result, 'Bookmark "pasted" (id 1) has more than one start boundary');
  });
});
