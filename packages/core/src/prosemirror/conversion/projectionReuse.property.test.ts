import { describe, expect, test } from "bun:test";
import fc from "fast-check";
import { EditorState } from "prosemirror-state";
import { registerSourceReplayDocument } from "@stll/docx-core/ops";

import { propertyConfig } from "../../../../../test/property-testing";
import {
  assignDocumentParagraphPropertySourceContract,
  PROSE_PARAGRAPH_SOURCE_TOKEN_ATTR,
} from "../../docx/paragraphPropertySource";
import type { BlockContent, Document, Paragraph, Table } from "../../types/document";
import { schema } from "../schema";
import { fromProseDoc } from "./fromProseDoc";
import { toProseDoc } from "./toProseDoc";
import { currentSourceProjection } from "./sourceProjection";

const sourceDocument = (content: BlockContent[]): Document => {
  const document: Document = { package: { document: { content } } };
  assignDocumentParagraphPropertySourceContract(document, "a".repeat(64));
  registerSourceReplayDocument(document);
  return document;
};

const paragraph = (hiddenFormatting: Paragraph["formatting"], text = "same"): Paragraph => ({
  type: "paragraph",
  ...(hiddenFormatting !== undefined ? { formatting: hiddenFormatting } : {}),
  content: [
    { type: "run", content: [{ type: "text", text }] },
    { type: "run", formatting: { bold: true }, content: [] },
    { type: "run", formatting: { italic: true }, content: [{ type: "text", text: "" }] },
    { type: "run", content: [{ type: "text", text }] },
  ],
});

describe("Document authority projection reuse", () => {
  test.each([0, 2, 4] as const)(
    "editing neighboring block %s retains distinct identical opaque sources",
    (editedIndex) => {
      fc.assert(
        fc.property(fc.string({ minLength: 1, maxLength: 16 }), (text) => {
          const first = {
            type: "preservedBlock",
            xml: '<e:opaque xmlns:e="urn:extension"/>',
          } satisfies BlockContent;
          const second = { ...first };
          const document = sourceDocument([
            paragraph(undefined, text),
            first,
            paragraph(undefined, text),
            second,
            paragraph(undefined, text),
          ]);
          const projected = toProseDoc(document);
          let position = 1;
          for (let index = 0; index < editedIndex; index += 1)
            position += projected.child(index).nodeSize;
          const edited = EditorState.create({ doc: projected }).tr.insertText("!", position).doc;
          const result = fromProseDoc(edited, document).package.document.content;
          expect(result.at(1)).toBe(first);
          expect(result.at(3)).toBe(second);
          expect(result.at(1)).not.toBe(result.at(3));
          expect(result.at(editedIndex)).not.toBe(
            document.package.document.content.at(editedIndex),
          );
          const moved = schema.node("doc", projected.attrs, [
            projected.child(0),
            projected.child(3),
            projected.child(2),
            projected.child(1),
            projected.child(4),
          ]);
          const reordered = fromProseDoc(moved, document).package.document.content;
          expect(reordered.at(1)).toBe(second);
          expect(reordered.at(3)).toBe(first);
          const duplicated = schema.node("doc", projected.attrs, [
            projected.child(0),
            projected.child(1),
            projected.child(2),
            projected.child(1),
            projected.child(4),
          ]);
          const duplicatedResult = fromProseDoc(duplicated, document).package.document.content;
          expect(duplicatedResult.at(1)).toBe(first);
          expect(duplicatedResult.at(3)).not.toBe(first);
          expect(duplicatedResult.at(3)).not.toBe(second);
          expect(duplicatedResult.at(3)).toEqual(
            fromProseDoc(duplicated, document, { reuse: "none" }).package.document.content.at(3),
          );
        }),
        propertyConfig(),
      );
    },
  );

  test.each(["copied", "changed", "foreign"] as const)(
    "a %s opaque editor node cannot borrow base ownership",
    (mode) => {
      const original = {
        type: "preservedBlock",
        xml: '<e:opaque xmlns:e="urn:extension"/>',
      } satisfies BlockContent;
      const document = sourceDocument([paragraph(undefined), original]);
      const projected = toProseDoc(document);
      const sourceNode = projected.child(1);
      const foreign = sourceDocument([paragraph(undefined), { ...original }]);
      const replacement =
        mode === "foreign"
          ? toProseDoc(foreign).child(1)
          : sourceNode.type.create(
              {
                ...sourceNode.attrs,
                ...(mode === "changed" ? { xml: '<e:changed xmlns:e="urn:extension"/>' } : {}),
              },
              sourceNode.content,
              sourceNode.marks,
            );
      const edited = schema.node("doc", projected.attrs, [projected.child(0), replacement]);
      const result = fromProseDoc(edited, document).package.document.content.at(1);
      expect(result).not.toBe(original);
      expect(result).not.toBe(foreign.package.document.content.at(1));
      expect(result).toEqual(
        fromProseDoc(edited, document, { reuse: "none" }).package.document.content.at(1),
      );
    },
  );
  test("unchanged projections retain every authored run, including empty runs", () => {
    fc.assert(
      fc.property(fc.string({ minLength: 1, maxLength: 16 }), (text) => {
        const original = paragraph(undefined, text);
        const document = sourceDocument([original]);
        const projected = toProseDoc(document);

        const unchanged = fromProseDoc(projected, document);
        expect(unchanged).not.toBe(document);
        expect(unchanged.package.document.content).not.toBe(document.package.document.content);
        expect(unchanged.package.document.content.at(0)).toBe(original);

        const rebuilt = fromProseDoc(projected, document, { reuse: "none" });
        expect(rebuilt.package.document.content).not.toBe(document.package.document.content);
        expect(rebuilt.package.document.content.at(0)).not.toBe(original);
        expect(rebuilt.package.document.content.at(0)).not.toEqual(original);
      }),
      propertyConfig(),
    );
  });

  test("numbered REF fields resolve after an edit to a warmed source projection", () => {
    const target: Paragraph = {
      type: "paragraph",
      formatting: { numPr: { kind: "reference", numId: 5, ilvl: 0 } },
      listRendering: {
        marker: "%1.",
        level: 0,
        numId: 5,
        abstractNumId: 7,
        isBullet: false,
        numFmt: "decimal",
        levelNumFmts: ["decimal"],
        levelStarts: [1],
      },
      content: [
        { type: "bookmarkStart", id: 1, name: "target" },
        { type: "run", content: [{ type: "text", text: "Target" }] },
        { type: "bookmarkEnd", id: 1 },
      ],
    };
    const reference: Paragraph = {
      type: "paragraph",
      content: [
        {
          type: "simpleField",
          instruction: " REF target \\w ",
          fieldType: "REF",
          content: [{ type: "run", content: [{ type: "text", text: "1" }] }],
        },
      ],
    };
    const document = sourceDocument([target, reference]);
    const projected = toProseDoc(document);

    fromProseDoc(projected, document);
    const inserted = schema.node(
      "paragraph",
      {
        ...projected.child(0).attrs,
        [PROSE_PARAGRAPH_SOURCE_TOKEN_ATTR]: null,
      },
      schema.text("Inserted"),
    );
    const edited = EditorState.create({ doc: projected }).tr.insert(0, inserted).doc;
    const resolved = fromProseDoc(edited, document);
    const savedReference = resolved.package.document.content.at(2);
    expect(resolved).not.toBe(document);
    expect(savedReference?.type).toBe("paragraph");
    if (savedReference?.type !== "paragraph") throw new TypeError("Expected REF paragraph");
    const field = savedReference.content.at(0);
    expect(field?.type).toBe("simpleField");
    if (field?.type !== "simpleField") throw new TypeError("Expected simple REF field");
    expect(field.content).toMatchObject([{ type: "run", content: [{ type: "text", text: "2" }] }]);
  });

  test("equal projections retain their own hidden content when reordered or deleted", () => {
    const first = paragraph(undefined);
    const second = paragraph(undefined);
    second.content.splice(1, 1, { type: "run", formatting: { noProof: true }, content: [] });
    const document = sourceDocument([first, second]);
    const projected = toProseDoc(document);
    expect(fromProseDoc(projected, document).package.document.content.at(0)).toBe(first);
    const firstNode = projected.child(0);
    const secondNode = projected.child(1);
    const moved = schema.node("doc", projected.attrs, [secondNode, firstNode]);
    const reordered = fromProseDoc(moved, document).package.document.content;

    expect(reordered.at(0)).toBe(second);
    expect(reordered.at(1)).toBe(first);
    const deleted = schema.node("doc", projected.attrs, [secondNode]);
    expect(fromProseDoc(deleted, document).package.document.content).toEqual([second]);
    expect(fromProseDoc(deleted, document).package.document.content.at(0)).toBe(second);
  });

  test("one edit rebuilds only the changed paragraph, without pPr or paragraph ids", () => {
    const first = paragraph(undefined, "before");
    const second = paragraph(undefined, "after");
    const document = sourceDocument([first, second]);
    const state = EditorState.create({ doc: toProseDoc(document) });
    const edited = state.tr.insertText("!", 2).doc;
    const result = fromProseDoc(edited, document).package.document.content;
    const rebuilt = fromProseDoc(edited, document, { reuse: "none" }).package.document.content;

    expect(result.at(0)).not.toBe(first);
    expect(result.at(0)).toEqual(rebuilt.at(0));
    expect(result.at(1)).toBe(second);
  });

  test("formatting changes cannot match on text alone", () => {
    const original = paragraph(undefined);
    const document = sourceDocument([original]);
    const state = EditorState.create({ doc: toProseDoc(document) });
    const edited = state.tr.addMark(1, 2, schema.mark("bold")).doc;
    const result = fromProseDoc(edited, document).package.document.content.at(0);

    expect(result).not.toBe(original);
    expect(result).toEqual(
      fromProseDoc(edited, document, { reuse: "none" }).package.document.content.at(0),
    );
  });

  test("a retained source token refers to the current Document after earlier edits", () => {
    const document = sourceDocument([paragraph(undefined)]);
    const state = EditorState.create({ doc: toProseDoc(document) });
    const edited = fromProseDoc(state.tr.insertText("!", 2).doc, document);
    const current = edited.package.document.content.at(0);

    expect(fromProseDoc(toProseDoc(edited), edited).package.document.content.at(0)).toBe(current);
    expect(current).not.toBe(document.package.document.content.at(0));
  });

  test("tracked base models reject mutations after warming source reuse", () => {
    // Vary model edits as well as editor edits: cached comparisons require
    // immutable authoritative records, including hidden authored content.
    fc.assert(
      fc.property(fc.string({ minLength: 1, maxLength: 16 }), (text) => {
        const original = paragraph(undefined, "old");
        const document = sourceDocument([original]);
        const projected = toProseDoc(document);
        const unchanged = fromProseDoc(projected, document);
        expect(unchanged).not.toBe(document);
        expect(unchanged.package.document.content).not.toBe(document.package.document.content);
        expect(unchanged.package.document.content.at(0)).toBe(original);
        expect(() =>
          original.content.push({ type: "run", content: [{ type: "text", text }] }),
        ).toThrow();
        const result = fromProseDoc(projected, document).package.document.content;

        expect(result.at(0)).toBe(original);
        expect(fromProseDoc(toProseDoc(document), document).package.document.content.at(0)).toBe(
          original,
        );
      }),
      propertyConfig(),
    );
  });

  test("mutable untracked models never cache projections or retain source runs", () => {
    fc.assert(
      fc.property(fc.string({ minLength: 1, maxLength: 16 }), (text) => {
        const original = paragraph(undefined, "old");
        const document: Document = { package: { document: { content: [original] } } };
        assignDocumentParagraphPropertySourceContract(document, "a".repeat(64));
        const projected = toProseDoc(document);
        expect(currentSourceProjection(document)).toBeUndefined();
        const unchanged = fromProseDoc(projected, document).package.document.content;
        expect(unchanged.at(0)).not.toBe(original);
        expect(unchanged).toEqual(
          fromProseDoc(projected, document, { reuse: "none" }).package.document.content,
        );
        original.content.push({ type: "run", content: [{ type: "text", text }] });
        expect(currentSourceProjection(document)).toBeUndefined();
        expect(fromProseDoc(projected, document).package.document.content).toEqual(
          fromProseDoc(projected, document, { reuse: "none" }).package.document.content,
        );
        const updated = toProseDoc(document);
        expect(currentSourceProjection(document)).toBeUndefined();
        expect(fromProseDoc(updated, document).package.document.content).toEqual(
          fromProseDoc(updated, document, { reuse: "none" }).package.document.content,
        );
      }),
      propertyConfig(),
    );
  });

  test("a projection with caller-supplied styles cannot authorize default-source reuse", () => {
    const original = paragraph(undefined);
    const document = sourceDocument([original]);
    document.package.styles = { docDefaults: { rPr: { bold: true } }, styles: [] };
    const projected = toProseDoc(document, { styles: { styles: [] } });
    const result = fromProseDoc(projected, document).package.document.content;

    expect(result.at(0)).not.toBe(original);
    expect(result).toEqual(
      fromProseDoc(projected, document, { reuse: "none" }).package.document.content,
    );
  });

  test("duplicate source tokens remain invalid even for identical projections", () => {
    const document = sourceDocument([paragraph(undefined)]);
    const projected = toProseDoc(document);
    expect(fromProseDoc(projected, document).package.document.content.at(0)).toBe(
      document.package.document.content.at(0),
    );
    const duplicate = schema.node("doc", projected.attrs, [projected.child(0), projected.child(0)]);

    expect(() => fromProseDoc(duplicate, document)).toThrow("more than one paragraph");
  });

  test("replacing a projected paragraph token cannot reuse its source object", () => {
    const original = paragraph(undefined);
    const document = sourceDocument([original]);
    const projected = toProseDoc(document);
    expect(fromProseDoc(projected, document).package.document.content.at(0)).toBe(original);
    const sourceNode = projected.child(0);
    const replacedNode = sourceNode.type.create(
      { ...sourceNode.attrs, [PROSE_PARAGRAPH_SOURCE_TOKEN_ATTR]: null },
      sourceNode.content,
      sourceNode.marks,
    );
    const replaced = schema.node("doc", projected.attrs, [replacedNode]);

    const result = fromProseDoc(replaced, document).package.document.content.at(0);
    expect(result).not.toBe(original);
    expect(result).toEqual(
      fromProseDoc(replaced, document, { reuse: "none" }).package.document.content.at(0),
    );
  });

  test("replacing a projected paragraph with changed formatting cannot reuse its source object", () => {
    const original = paragraph(undefined);
    const document = sourceDocument([original]);
    const projected = toProseDoc(document);
    expect(fromProseDoc(projected, document).package.document.content.at(0)).toBe(original);
    const edited = EditorState.create({ doc: projected }).tr.addMark(1, 2, schema.mark("bold")).doc;

    const result = fromProseDoc(edited, document).package.document.content.at(0);
    expect(result).not.toBe(original);
    expect(result).toEqual(
      fromProseDoc(edited, document, { reuse: "none" }).package.document.content.at(0),
    );
  });

  test("equal text without source identities is rebuilt", () => {
    const original = paragraph(undefined);
    const document: Document = { package: { document: { content: [original] } } };

    expect(fromProseDoc(toProseDoc(document), document).package.document.content.at(0)).not.toBe(
      original,
    );
  });

  test("unchanged source projections cannot bypass suggestion stripping", () => {
    const info = {
      id: 1,
      author: "Reviewer",
      date: "2026-10-01T00:00:00Z",
      provenance: "suggested",
    };
    const original: Paragraph = {
      type: "paragraph",
      formatting: { alignment: "center" },
      propertyChanges: [
        {
          type: "paragraphPropertyChange",
          info,
          previousFormatting: { alignment: "left" },
          currentFormatting: { alignment: "center" },
        },
      ],
      content: [{ type: "run", content: [{ type: "text", text: "keep" }] }],
    };
    const document = sourceDocument([original]);
    const projected = toProseDoc(document);
    const result = fromProseDoc(projected, document).package.document.content.at(0);

    expect(result).not.toBe(original);
    expect(result).toEqual(
      fromProseDoc(projected, document, { reuse: "none" }).package.document.content.at(0),
    );
    expect(result).toMatchObject({ formatting: { alignment: "left" } });
    expect(JSON.stringify(result)).not.toContain("suggested");
  });

  test("removing a projected page break cannot restore the source break", () => {
    const original = paragraph(undefined, "before");
    original.content.push({ type: "run", content: [{ type: "break", breakType: "page" }] });
    const document = sourceDocument([original]);
    const projected = toProseDoc(document);
    const children = Array.from({ length: projected.childCount }, (_, index) =>
      projected.child(index),
    ).filter((node) => node.type.name !== "pageBreak");
    const edited = schema.node("doc", projected.attrs, children);

    expect(fromProseDoc(edited, document).package.document.content.at(0)).not.toBe(original);
    expect(fromProseDoc(edited, document).package.document.content).toEqual(
      fromProseDoc(edited, document, { reuse: "none" }).package.document.content,
    );
  });

  test("unchanged containers retain their authoritative identity, edited containers retain children", () => {
    const first = paragraph(undefined, "first");
    const second = paragraph(undefined, "second");
    const table: Table = {
      type: "table",
      rows: [
        {
          type: "tableRow",
          cells: [{ type: "tableCell", content: [first, second] }],
        },
      ],
    };
    const container = {
      type: "blockSdt",
      properties: { sdtType: "richText" },
      content: [table],
    } as const satisfies BlockContent;
    const document = sourceDocument([container]);
    const projected = toProseDoc(document);

    expect(fromProseDoc(projected, document).package.document.content.at(0)).toBe(container);
    const state = EditorState.create({ doc: projected });
    let firstTextPosition: number | undefined;
    projected.descendants((node, position) => {
      if (node.isText && firstTextPosition === undefined) {
        firstTextPosition = position;
      }
    });
    if (firstTextPosition === undefined) {
      throw new Error("Expected text inside the table");
    }
    const edited = fromProseDoc(state.tr.insertText("!", firstTextPosition).doc, document);
    const savedContainer = edited.package.document.content.at(0);
    if (savedContainer?.type !== "blockSdt") {
      throw new Error("Expected the content control");
    }
    const savedTable = savedContainer.content.at(0);
    if (savedTable?.type !== "table") {
      throw new Error("Expected the table");
    }
    expect(savedContainer).not.toBe(container);
    expect(savedTable).not.toBe(table);
    const savedParagraphs = savedTable.rows.at(0)?.cells.at(0)?.content;
    expect(savedParagraphs?.at(0)).not.toBe(first);
    expect(savedParagraphs?.at(1)).toBe(second);
  });
});
