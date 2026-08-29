import { describe, expect, test } from "bun:test";

import { schema } from "./schema";
import { assertValidProseMirrorDocument, validateProseMirrorDocument } from "./validation";

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
      message: "Bookmark id 1 has no matching end boundary.",
    },
    {
      name: "lone end",
      content: [schema.node("bookmarkBoundary", { type: "end", id: 1 })],
      message: "Bookmark id 1 has no open start boundary.",
    },
    {
      name: "duplicate starts",
      content: [
        schema.node("bookmarkBoundary", { type: "start", id: 1, name: "one" }),
        schema.node("bookmarkBoundary", { type: "end", id: 1 }),
        schema.node("bookmarkBoundary", { type: "start", id: 1, name: "duplicate" }),
        schema.node("bookmarkBoundary", { type: "end", id: 1 }),
      ],
      message: "Bookmark id 1 has more than one start boundary.",
    },
  ])("rejects $name bookmark boundaries", ({ content, message }) => {
    const doc = schema.node("doc", null, [schema.node("paragraph", null, content)]);

    const result = validateProseMirrorDocument(doc);

    expect(result.valid).toBe(false);
    expect(result.issues.map((issue) => issue.message)).toContain(message);
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

  test("rejects a tracked boundary placement the document model cannot serialize", () => {
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

    expect(result.valid).toBe(false);
    expect(result.issues.map((issue) => issue.message)).toContain(
      "Bookmark boundaries inside tracked changes require a hyperlink serialization parent.",
    );
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
      "Structured simple fields require hyperlink content.",
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
    expect(result.issues.map((issue) => issue.message)).toContain(
      "Bookmark id 1 has more than one start boundary.",
    );
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
    expect(result.issues.map((issue) => issue.message)).toContain(
      "Bookmark id 1 has more than one start boundary.",
    );
  });
});
