import { describe, expect, test } from "bun:test";
import type { Node as PMNode } from "prosemirror-model";
import { EditorState } from "prosemirror-state";
import { history, undo } from "prosemirror-history";

import type { Document, Paragraph, StyleDefinitions } from "../../types/document";
import { fromProseDoc } from "../conversion/fromProseDoc";
import { toProseDoc } from "../conversion/toProseDoc";
import { schema } from "../schema";
import { createDocumentStylesPlugin } from "./documentStyles";

const STYLES: StyleDefinitions = {
  docDefaults: {
    pPr: { spaceAfter: 160, lineSpacing: 259, lineSpacingRule: "auto" },
    rPr: { fontSize: 22 },
  },
  styles: [
    { styleId: "Normal", type: "paragraph", default: true },
    {
      styleId: "Heading1",
      type: "paragraph",
      basedOn: "Normal",
      pPr: { spaceBefore: 240, keepNext: true },
      rPr: { bold: true, fontSize: 32 },
    },
    {
      styleId: "TableGrid",
      type: "table",
      pPr: { spaceAfter: 0, lineSpacing: 240, lineSpacingRule: "auto" },
      tblStylePr: [
        { type: "firstRow", pPr: { spaceBefore: 120 } },
        { type: "band2Horz", pPr: { spaceAfter: 60 } },
      ],
    },
  ],
};

const loadedParagraph = (text: string): Paragraph => ({
  type: "paragraph",
  content: [{ type: "run", content: [{ type: "text", text }] }],
});

const loadState = (): EditorState => {
  const document: Document = {
    package: { document: { content: [loadedParagraph("Loaded")] }, styles: STYLES },
  };
  return EditorState.create({
    doc: toProseDoc(document),
    plugins: [history(), createDocumentStylesPlugin(STYLES)],
  });
};

const paragraph = (text: string, attrs: Record<string, unknown> = {}): PMNode =>
  schema.nodes["paragraph"]!.create(attrs, schema.text(text));

const insertAtEnd = (state: EditorState, ...nodes: PMNode[]): EditorState =>
  state.apply(state.tr.insert(state.doc.content.size, nodes));

const paragraphs = (doc: PMNode): PMNode[] => {
  const found: PMNode[] = [];
  doc.descendants((node) => {
    if (node.type.name === "paragraph") {
      found.push(node);
    }
    return true;
  });
  return found;
};

const byText = (doc: PMNode, text: string): PMNode => {
  const node = paragraphs(doc).find((candidate) => candidate.textContent === text);
  if (!node) {
    throw new Error(`No paragraph "${text}"`);
  }
  return node;
};

const savedFormatting = (doc: PMNode, text: string): Paragraph["formatting"] => {
  const saved = fromProseDoc(doc, { package: { document: { content: [] } } });
  const block = saved.package.document.content.find(
    (candidate): candidate is Paragraph =>
      candidate.type === "paragraph" &&
      candidate.content.some(
        (item) =>
          item.type === "run" &&
          item.content.some((content) => content.type === "text" && content.text === text),
      ),
  );
  return block?.formatting;
};

describe("paragraphs an edit creates resolve their style cascade", () => {
  test("a new paragraph paints the document defaults a loaded one does", () => {
    const loaded = loadState();
    const after = insertAtEnd(loaded, paragraph("Pasted"));
    const reference = byText(after.doc, "Loaded").attrs;
    const created = byText(after.doc, "Pasted").attrs;

    for (const key of ["spaceAfter", "lineSpacing", "lineSpacingRule", "defaultTextFormatting"]) {
      expect(created[key]).toEqual(reference[key]);
    }
    expect(created["spaceAfter"]).toBe(160);
    // The inherited values stay out of the paragraph's own `w:pPr`.
    expect(savedFormatting(after.doc, "Pasted")?.spaceAfter).toBeUndefined();
    expect(savedFormatting(after.doc, "Pasted")?.lineSpacing).toBeUndefined();
  });

  test("a new paragraph reads its own style chain and keeps what it states", () => {
    const after = insertAtEnd(
      loadState(),
      paragraph("Heading", { styleId: "Heading1" }),
      paragraph("Tight", { spaceAfter: 0 }),
    );
    const heading = byText(after.doc, "Heading").attrs;
    expect(heading["spaceBefore"]).toBe(240);
    expect(heading["spaceAfter"]).toBe(160);
    expect(heading["keepNext"]).toBe(true);
    expect(heading["defaultTextFormatting"]).toMatchObject({ bold: true, fontSize: 32 });
    expect(savedFormatting(after.doc, "Heading")?.spaceBefore).toBeUndefined();
    expect(savedFormatting(after.doc, "Heading")?.keepNext).toBeUndefined();

    const tight = byText(after.doc, "Tight").attrs;
    expect(tight["spaceAfter"]).toBe(0);
    expect(tight["lineSpacing"]).toBe(259);
    expect(savedFormatting(after.doc, "Tight")?.spaceAfter).toBe(0);
  });

  test("a new cell paragraph layers the table style over the document defaults", () => {
    const cell = schema.nodes["tableCell"]!.create(null, paragraph("Cell"));
    const table = schema.nodes["table"]!.create(
      { styleId: "TableGrid" },
      schema.nodes["tableRow"]!.create(null, cell),
    );
    const created = byText(insertAtEnd(loadState(), table).doc, "Cell").attrs;
    expect(created["spaceAfter"]).toBe(0);
    expect(created["lineSpacing"]).toBe(240);
  });

  test("a new cell paragraph reads its cell's table-style regions as a loaded one does", () => {
    const loadedCell = (text: string) => ({
      type: "tableCell" as const,
      content: [loadedParagraph(text)],
    });
    const document: Document = {
      package: {
        document: {
          content: [
            {
              type: "table",
              formatting: { styleId: "TableGrid", look: { firstRow: true } },
              columnWidths: [2000],
              rows: ["Header", "First", "Second"].map((text) => ({
                type: "tableRow" as const,
                cells: [loadedCell(text)],
              })),
            },
          ],
        },
        styles: STYLES,
      },
    };
    const loaded = EditorState.create({
      doc: toProseDoc(document),
      plugins: [createDocumentStylesPlugin(STYLES)],
    });
    // Replace each loaded cell paragraph with a fresh one, as a cut or a
    // structural command would.
    let edited = loaded;
    for (const text of ["Header", "First", "Second"]) {
      let found = -1;
      edited.doc.descendants((node, pos) => {
        if (node.type.name === "paragraph" && node.textContent === text) {
          found = pos;
        }
        return found < 0;
      });
      const node = edited.doc.nodeAt(found)!;
      edited = edited.apply(
        edited.tr.replaceWith(found, found + node.nodeSize, paragraph(`New ${text}`)),
      );
    }
    const keys = ["spaceBefore", "spaceAfter", "lineSpacing", "lineSpacingRule"];
    for (const text of ["Header", "First", "Second"]) {
      const reference = byText(loaded.doc, text).attrs;
      const created = byText(edited.doc, `New ${text}`).attrs;
      for (const key of keys) {
        expect(created[key]).toEqual(reference[key]);
      }
    }
    expect(byText(edited.doc, "New Header").attrs["spaceBefore"]).toBe(120);
    expect(byText(edited.doc, "New First").attrs["spaceAfter"]).toBe(0);
    expect(byText(edited.doc, "New Second").attrs["spaceAfter"]).toBe(60);
  });

  test("the resolution undoes with the edit that needed it", () => {
    const loaded = loadState();
    const after = insertAtEnd(loaded, paragraph("Pasted"));
    let undone = after;
    undo(after, (tr) => {
      undone = after.apply(tr);
    });
    expect(undone.doc.eq(loaded.doc)).toBe(true);
  });

  test("an edit inside a loaded paragraph leaves its attrs alone", () => {
    const loaded = loadState();
    const after = loaded.apply(loaded.tr.insertText("!", 1));
    expect(after.doc.firstChild?.attrs).toEqual(loaded.doc.firstChild?.attrs);
  });
});
