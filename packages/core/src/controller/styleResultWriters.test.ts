/**
 * A writer that builds something it may change out of style results or of a
 * paragraph's style-derived attributes takes its own copy first
 * (`writableCopy`). Each test here deep-freezes what a writer reads, runs the
 * writer, and checks that it neither throws on the frozen input nor returns an
 * object it shares with that input. The save copies only the values a node
 * states itself, so its copies grow with the authored values, never with the
 * document.
 */

import { describe, expect, spyOn, test } from "bun:test";
import type { Node as PMNode } from "prosemirror-model";

import { deepFreeze } from "../__tests__/deepFreeze";
import { importReferencedStyleDefinitions } from "../compare/style-resources";
import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { buildRunFormattingOverrideAttrs } from "../prosemirror/extensions/marks/RunFormattingOverrideExtension";
import { schema } from "../prosemirror/schema";
import type { Document, Paragraph, StyleDefinitions } from "../types/document";
import * as documentModel from "../types/documentModel";
import { importClipboardStyles } from "./canonicalClipboardResources";

/** Every object reachable from `value`. */
const reachable = (value: unknown, seen = new Set<object>()): Set<object> => {
  if (typeof value !== "object" || value === null || seen.has(value))
    return seen;
  seen.add(value);
  for (const entry of Object.values(value)) reachable(entry, seen);
  return seen;
};

/** The objects `result` shares with `input`. */
const sharedObjects = (input: unknown, result: unknown): object[] => {
  const inputs = reachable(input);
  return [...reachable(result)].filter((object) => inputs.has(object));
};

const BORDER = {
  style: "single",
  size: 8,
  color: { rgb: "112233" },
} as const;

const sourceStyles = (): StyleDefinitions => ({
  docDefaults: {
    rPr: {
      bold: true,
      underline: { style: "single", color: { rgb: "445566" } },
    },
    pPr: { spaceAfter: 120, tabs: [{ position: 720, alignment: "left" }] },
  },
  styles: [
    {
      styleId: "Body",
      type: "paragraph",
      default: true,
      name: "Normal",
      pPr: {
        borders: { top: { ...BORDER } },
        shading: { fill: { rgb: "EEEEEE" }, pattern: "clear" },
      },
      rPr: { color: { rgb: "778899" }, fontFamily: { ascii: "Arial" } },
    },
    {
      styleId: "Strong",
      type: "character",
      rPr: { italic: true, shading: { fill: { rgb: "DDDDDD" } } },
    },
  ],
});

const documentWith = (styles: StyleDefinitions): Document => ({
  package: { document: { content: [] }, styles },
});

describe("style result writers copy first", () => {
  test("pasted paragraphs own the formatting the source styles resolve to", () => {
    const paragraphs: Paragraph[] = [
      { type: "paragraph", content: [] },
      {
        type: "paragraph",
        formatting: { styleId: "Body", runProperties: { styleId: "Strong" } },
        content: [
          {
            type: "run",
            formatting: { styleId: "Strong" },
            content: [{ type: "text", text: "pasted" }],
          },
        ],
      },
    ];
    const plain = importClipboardStyles({
      source: documentWith(sourceStyles()),
      destination: documentWith({ styles: [] }),
      paragraphs: structuredClone(paragraphs),
    });
    const source = deepFreeze(documentWith(sourceStyles()));
    const destination = deepFreeze(documentWith({ styles: [] }));
    const frozen = importClipboardStyles({
      source,
      destination,
      paragraphs: deepFreeze(structuredClone(paragraphs)),
    });
    expect(frozen).toEqual(plain);
    if (frozen.isErr()) throw new Error(frozen.error.message);
    expect(sharedObjects(source, frozen.value.paragraphs)).toEqual([]);
    expect(sharedObjects(destination, frozen.value.paragraphs)).toEqual([]);
    // Two paragraphs of one style get two records, not the engine's one.
    const [first, second] = frozen.value.paragraphs;
    expect(sharedObjects(first, second)).toEqual([]);
  });

  test("a materialized style owns the values its source resolves to", () => {
    const run = (styles: StyleDefinitions) =>
      importReferencedStyleDefinitions({
        sourceStyles: styles,
        destinationStyles: { styles: [] },
        sourceTheme: undefined,
        destinationTheme: undefined,
        referencedStyleIds: ["Body"],
        materializeDefaultParagraphStyle: true,
      });
    const plain = run(sourceStyles());
    const source = deepFreeze(sourceStyles());
    const frozen = run(source);
    expect(frozen).toEqual(plain);
    if (frozen.status === "unalignable") throw new Error(frozen.detail);
    const imported = frozen.styles?.styles.filter(({ styleId }) =>
      frozen.importedStyleIds.includes(styleId),
    );
    expect(imported?.length).toBeGreaterThan(0);
    expect(sharedObjects(source, imported)).toEqual([]);
  });

  test("a saved paragraph and cell own the borders, shading and tabs of their attrs", () => {
    const paragraphAttrs = deepFreeze({
      borders: { top: { ...BORDER } },
      shading: { fill: { rgb: "ABCDEF" }, pattern: "clear" },
      tabs: [{ position: 1440, alignment: "right", leader: "dot" }],
    });
    const cellAttrs = deepFreeze({ borders: { bottom: { ...BORDER } } });
    const doc = schema.node("doc", null, [
      schema.node("paragraph", paragraphAttrs, [schema.text("bordered")]),
      schema.node("table", null, [
        schema.node("tableRow", null, [
          schema.node("tableCell", cellAttrs, [schema.node("paragraph")]),
        ]),
      ]),
    ]);
    const saved = fromProseDoc(doc, undefined, {
      stylesheetSource: { type: "package" },
    });
    const [paragraph, table] = saved.package.document.content;
    if (paragraph?.type !== "paragraph" || table?.type !== "table")
      throw new Error("The saved body lost its blocks.");
    expect(paragraph.formatting).toMatchObject(paragraphAttrs);
    expect(sharedObjects(paragraphAttrs, paragraph.formatting)).toEqual([]);
    const cellFormatting = table.rows[0]?.cells[0]?.formatting;
    expect(cellFormatting?.borders).toEqual(cellAttrs.borders);
    expect(sharedObjects(cellAttrs, cellFormatting)).toEqual([]);
  });

  test("the save copies once per authored value, not per document node", () => {
    const tableOf = (rows: number, bordered: (index: number) => boolean) =>
      schema.node(
        "table",
        null,
        Array.from({ length: rows }, (_, row) =>
          schema.node(
            "tableRow",
            null,
            Array.from({ length: 4 }, (_cell, column) =>
              schema.node(
                "tableCell",
                bordered(row * 4 + column)
                  ? { borders: { top: { ...BORDER } } }
                  : null,
                [schema.node("paragraph", null, [schema.text("cell")])],
              ),
            ),
          ),
        ),
      );
    const copiesFor = (doc: PMNode): number => {
      const copies = spyOn(documentModel, "writableCopy");
      try {
        fromProseDoc(doc, undefined, { stylesheetSource: { type: "package" } });
        return copies.mock.calls.length;
      } finally {
        copies.mockRestore();
      }
    };
    const plain = copiesFor(
      schema.node("doc", null, [tableOf(50, () => false)]),
    );
    const someBordered = copiesFor(
      schema.node("doc", null, [tableOf(50, (index) => index % 10 === 0)]),
    );
    const moreRows = copiesFor(
      schema.node("doc", null, [tableOf(100, (index) => index < 20)]),
    );
    // A cell that states no borders costs no copy; one that does costs one.
    expect(plain).toBe(0);
    expect(someBordered).toBe(20);
    expect(moreRows).toBe(20);
  });

  test("a cleared shading override owns its shading", () => {
    const formatting = deepFreeze({
      shading: {
        pattern: "nil",
        color: { rgb: "010203" },
        preservedAttributes: [{ name: "w:themeFill", value: "accent1" }],
      },
    } as const);
    const attrs = buildRunFormattingOverrideAttrs(formatting);
    expect(attrs?.shading).toEqual(formatting.shading);
    expect(sharedObjects(formatting.shading, attrs?.shading)).toEqual([]);
  });
});
