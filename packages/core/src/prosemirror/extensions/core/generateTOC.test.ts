import { describe, expect, test } from "bun:test";
import type { Node as PMNode } from "prosemirror-model";
import { EditorState, TextSelection } from "prosemirror-state";
import type { Transaction } from "prosemirror-state";

import { createDocumentStylesPlugin } from "../../plugins/documentStyles";
import { schema, singletonManager } from "../../schema";
import type { StyleDefinitions } from "../../../types/document";

/**
 * A Czech Word's heading styles: the ids carry the localized, accent-stripped
 * UI name, and only `w:name` says what they are.
 */
const LOCALIZED_STYLES: StyleDefinitions = {
  styles: [
    { styleId: "Nadpis1", type: "paragraph", name: "heading 1" },
    { styleId: "Nadpis2", type: "paragraph", name: "heading 2" },
    { styleId: "Obsah1", type: "paragraph", name: "toc 1" },
    { styleId: "Obsah2", type: "paragraph", name: "toc 2" },
  ],
};

const docWithHeadings = (styleIds: readonly [string, string]): PMNode =>
  schema.node("doc", null, [
    schema.node("paragraph", { styleId: styleIds[0] }, [schema.text("Introduction")]),
    schema.node("paragraph", {}, [schema.text("Body text.")]),
    schema.node("paragraph", { styleId: styleIds[1] }, [schema.text("Background")]),
  ]);

/** An English Word's, with the outline levels Word's built-ins carry. */
const ENGLISH_STYLES: StyleDefinitions = {
  styles: [
    { styleId: "Heading1", type: "paragraph", name: "heading 1", pPr: { outlineLevel: 0 } },
    { styleId: "Heading2", type: "paragraph", name: "heading 2", pPr: { outlineLevel: 1 } },
    { styleId: "TOC1", type: "paragraph", name: "toc 1" },
    { styleId: "TOC2", type: "paragraph", name: "toc 2" },
  ],
};

const runGenerateTOC = (doc: PMNode, styles: StyleDefinitions): PMNode => {
  const generateTOC = singletonManager.getCommands()["generateTOC"];
  if (!generateTOC) {
    throw new Error("generateTOC command not registered");
  }
  let state = EditorState.create({
    schema,
    doc,
    plugins: [createDocumentStylesPlugin(styles)],
  });
  state = state.apply(state.tr.setSelection(TextSelection.atStart(state.doc)));

  let captured: Transaction | undefined;
  const ok = generateTOC({ title: "Table of Contents" })(state, (tr) => {
    captured = tr;
  });
  expect(ok).toBe(true);
  if (!captured) {
    throw new Error("generateTOC did not dispatch a transaction");
  }
  return captured.doc;
};

const TOC_ENTRY_STYLE_IDS = new Set(["TOC1", "TOC2", "TOC3", "Obsah1", "Obsah2"]);

const tocEntryParagraphs = (doc: PMNode): PMNode[] => {
  const entries: PMNode[] = [];
  doc.descendants((node) => {
    const styleId = node.attrs["styleId"];
    if (
      node.type.name === "paragraph" &&
      typeof styleId === "string" &&
      TOC_ENTRY_STYLE_IDS.has(styleId)
    ) {
      entries.push(node);
    }
  });
  return entries;
};

const HEADING_STYLE_IDS = new Set(["Heading1", "Heading2", "Nadpis1", "Nadpis2"]);

describe("generateTOC", () => {
  test("creates one entry per heading with a PAGEREF field and a dot-leader right tab", () => {
    const result = runGenerateTOC(docWithHeadings(["Heading1", "Heading2"]), ENGLISH_STYLES);
    const entries = tocEntryParagraphs(result);

    expect(entries).toHaveLength(2);

    for (const entry of entries) {
      let pagerefInstruction: string | undefined;
      let hasTab = false;
      entry.descendants((node) => {
        if (node.type.name === "field" && node.attrs["fieldType"] === "PAGEREF") {
          pagerefInstruction = node.attrs["instruction"] as string;
        }
        if (node.type.name === "tab") {
          hasTab = true;
        }
      });

      // PAGEREF points at a generated heading bookmark.
      expect(pagerefInstruction).toMatch(/^PAGEREF _Toc\d+ \\h$/u);
      expect(hasTab).toBe(true);

      const tabs = entry.attrs["tabs"] as { alignment?: string; leader?: string }[] | null;
      expect(tabs?.some((t) => t.alignment === "right" && t.leader === "dot")).toBe(true);
    }
  });

  test("each entry's PAGEREF targets a bookmark anchored on a heading", () => {
    const result = runGenerateTOC(docWithHeadings(["Heading1", "Heading2"]), ENGLISH_STYLES);

    // Bookmark names anchored on heading paragraphs.
    const headingBookmarks = new Set<string>();
    result.descendants((node) => {
      if (node.type.name !== "paragraph") {
        return;
      }
      const styleId = node.attrs["styleId"];
      if (typeof styleId === "string" && HEADING_STYLE_IDS.has(styleId)) {
        const bookmarks = node.attrs["bookmarks"] as { name: string }[] | undefined;
        for (const b of bookmarks ?? []) {
          headingBookmarks.add(b.name);
        }
      }
    });

    const entries = tocEntryParagraphs(result);
    expect(entries).toHaveLength(2);
    for (const entry of entries) {
      let target: string | undefined;
      entry.descendants((node) => {
        if (node.type.name === "field" && node.attrs["fieldType"] === "PAGEREF") {
          target = /^PAGEREF (?<bookmark>\S+) /u.exec(node.attrs["instruction"] as string)
            ?.groups?.["bookmark"];
        }
      });
      expect(target).toBeDefined();
      // The referenced bookmark really exists on a heading, so the page map
      // resolves it at paint.
      expect(headingBookmarks.has(target as string)).toBe(true);
    }
  });

  test("collects headings a localized Word wrote, which carry no English id", () => {
    // `Nadpis1`/`Nadpis2` are what a Czech Word writes; the styles carry no
    // outline level, so only `w:name` identifies them.
    const result = runGenerateTOC(docWithHeadings(["Nadpis1", "Nadpis2"]), LOCALIZED_STYLES);
    expect(tocEntryParagraphs(result)).toHaveLength(2);
  });

  test("uses the caller's title and the document's own TOC styles", () => {
    const result = runGenerateTOC(docWithHeadings(["Nadpis1", "Nadpis2"]), LOCALIZED_STYLES);
    const texts: string[] = [];
    result.descendants((node) => {
      if (node.type.name === "paragraph") {
        texts.push(node.textContent);
      }
    });
    expect(texts).toContain("Table of Contents");
    // The document calls its TOC entry styles `Obsah1`/`Obsah2`; writing
    // `TOC1` would name a style it does not define.
    expect(tocEntryParagraphs(result).map((node) => node.attrs["styleId"])).toEqual([
      "Obsah1",
      "Obsah2",
    ]);
  });

  test("writes no style id when the document defines no TOC styles", () => {
    const bare: StyleDefinitions = {
      styles: [{ styleId: "Nadpis1", type: "paragraph", name: "heading 1" }],
    };
    const result = runGenerateTOC(docWithHeadings(["Nadpis1", "Nadpis1"]), bare);
    const styleIds: unknown[] = [];
    result.descendants((node) => {
      if (node.type.name === "paragraph") {
        styleIds.push(node.attrs["styleId"]);
      }
    });
    // Only the two source headings carry a style; the generated paragraphs
    // carry none rather than a dangling `TOCHeading`/`TOC1`.
    expect(styleIds.filter((id) => id === "TOCHeading" || id === "TOC1")).toEqual([]);
  });
});
