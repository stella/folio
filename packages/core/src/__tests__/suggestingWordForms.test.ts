/**
 * What suggesting mode records for edits that remove or add paragraph breaks,
 * checked against the revisions a reference implementation writes for the
 * same edits and against what accepting and rejecting them leaves.
 */

import { describe, expect, test } from "bun:test";
import { readFileSync } from "node:fs";
import path from "node:path";

import JSZip from "jszip";
import { undo } from "prosemirror-history";
import { AllSelection, TextSelection, type EditorState } from "prosemirror-state";

import { expectParagraphAttrs } from "../prosemirror/attrs";
import type { Document } from "../types/document";
import { acceptAIEditRevision } from "../prosemirror/commands/comments";
import { deleteSelectionAsSuggestion } from "../prosemirror/plugins/suggestionMode";
import {
  createHarnessState,
  HeadlessEditorView,
  parseShapeDocument,
  placeSelection,
  resolveAllChanges,
  saveHarnessState,
  textblocks,
} from "./editorHarness";
import type { EditorMode } from "./editorHarness";

const reference = JSON.parse(
  readFileSync(
    path.join(import.meta.dir, "__fixtures__", "revision-resolution-reference.json"),
    "utf8",
  ),
) as { package: { styles: string; numbering: string; settings: string; sectPr: string } };

const MAIN = "application/vnd.openxmlformats-officedocument.wordprocessingml";
const RELATIONSHIPS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NAMESPACES =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  `xmlns:r="${RELATIONSHIPS}" ` +
  'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" ' +
  'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w14"';

const paragraph = (id: string, text: string, properties = "") =>
  `<w:p w14:paraId="${id}" w14:textId="${id}">${properties ? `<w:pPr>${properties}</w:pPr>` : ""}<w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;

const table = () =>
  `<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="4680"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="4680" w:type="dxa"/></w:tcPr>${paragraph("66000001", "Cell.")}</w:tc></w:tr></w:tbl>`;

const packageOf = async (body: string): Promise<Uint8Array> => {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${MAIN}.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="${MAIN}.styles+xml"/><Override PartName="/word/numbering.xml" ContentType="${MAIN}.numbering+xml"/><Override PartName="/word/settings.xml" ContentType="${MAIN}.settings+xml"/></Types>`,
  );
  zip.file(
    "_rels/.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${RELATIONSHIPS}/officeDocument" Target="word/document.xml"/></Relationships>`,
  );
  zip.file(
    "word/_rels/document.xml.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${RELATIONSHIPS}/styles" Target="styles.xml"/><Relationship Id="rId2" Type="${RELATIONSHIPS}/numbering" Target="numbering.xml"/><Relationship Id="rId3" Type="${RELATIONSHIPS}/settings" Target="settings.xml"/></Relationships>`,
  );
  zip.file(
    "word/document.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NAMESPACES}><w:body>${body}${reference.package.sectPr}</w:body></w:document>`,
  );
  zip.file("word/styles.xml", reference.package.styles);
  zip.file("word/numbering.xml", reference.package.numbering);
  zip.file("word/settings.xml", reference.package.settings);
  return zip.generateAsync({ type: "uint8array" });
};

const open = async (body: string, mode: EditorMode = "suggesting") => {
  const document = await parseShapeDocument(await packageOf(body));
  return { document, state: createHarnessState(document, mode) };
};

/** Text, paragraph id, alignment, style and mark revision of every top-level paragraph. */
const paragraphs = (doc: import("prosemirror-model").Node) =>
  textblocks(doc).map(({ node }) => {
    const attrs = expectParagraphAttrs(node);
    return {
      text: node.textContent,
      paraId: attrs.paraId ?? null,
      alignment: attrs.alignment ?? null,
      styleId: attrs.styleId ?? null,
      mark: attrs.pPrMark?.kind ?? null,
      propertyChange: (attrs._propertyChanges ?? []).length > 0,
    };
  });

/** The paragraphs as read, whichever paragraph ids they carry. */
const withoutIds = (entries: ReturnType<typeof paragraphs>) =>
  entries.map((entry) => Object.assign(entry, { paraId: expect.any(String) as unknown as string }));

const INTRO = paragraph("1A000001", "Intro paragraph.");
const CENTERED = paragraph("2A000002", "Alpha text.", '<w:jc w:val="center"/>');
const RIGHT = paragraph("3B000003", "Bravo text.", '<w:jc w:val="right"/>');
const TAIL = paragraph("5E000005", "Tail paragraph.");

describe("suggesting mode writes the revisions the reference writes", () => {
  for (const [key, placement, focus] of [
    ["Backspace", "caret-start", "Bravo text."],
    ["Mod-Backspace", "caret-start", "Bravo text."],
    ["Delete", "caret-end", "Alpha text."],
    ["Mod-Delete", "caret-end", "Alpha text."],
  ] as const) {
    test(`${key} at the break deletes the mark and gives the next paragraph the first one's properties`, async () => {
      const { state } = await open(INTRO + CENTERED + RIGHT + TAIL);
      const view = new HeadlessEditorView(placeSelection(state, focus, placement) ?? state);
      view.pressKey(key);
      expect(paragraphs(view.state.doc).slice(1, 3)).toEqual([
        {
          text: "Alpha text.",
          paraId: "2A000002",
          alignment: "center",
          styleId: null,
          mark: "del",
          propertyChange: false,
        },
        {
          text: "Bravo text.",
          paraId: "3B000003",
          alignment: "center",
          styleId: null,
          mark: null,
          propertyChange: true,
        },
      ]);
      expect(paragraphs(resolveAllChanges(view.state, "accept").doc).slice(1, 2)).toEqual([
        {
          text: "Alpha text.Bravo text.",
          paraId: "3B000003",
          alignment: "center",
          styleId: null,
          mark: null,
          propertyChange: false,
        },
      ]);
      expect(paragraphs(resolveAllChanges(view.state, "reject").doc)).toEqual(
        paragraphs(state.doc),
      );
    });
  }

  test("a cut across paragraphs gives the paragraph left the first one's properties", async () => {
    const { state } = await open(INTRO + CENTERED + RIGHT + TAIL);
    const placed = placeSelection(state, "Alpha text.", "cross-paragraph") ?? state;
    const view = new HeadlessEditorView(placed);
    expect(deleteSelectionAsSuggestion(view.state, view.dispatch)).toBe(true);
    const accepted = paragraphs(resolveAllChanges(view.state, "accept").doc);
    expect(accepted[1]).toMatchObject({ paraId: "3B000003", alignment: "center" });
    expect(paragraphs(resolveAllChanges(view.state, "reject").doc)).toEqual(paragraphs(state.doc));
  });

  test("typing over select-all replaces the text into the last paragraph", async () => {
    const body =
      paragraph("2A000002", "Alpha text.", '<w:pStyle w:val="Heading1"/><w:jc w:val="center"/>') +
      paragraph("3B000003", "Bravo text.") +
      paragraph("4C000004", "Charlie text.", '<w:pStyle w:val="BlockQuote"/><w:jc w:val="right"/>');
    for (const mode of ["suggesting", "editing"] as const) {
      const { state } = await open(body, mode);
      const view = new HeadlessEditorView(
        state.apply(state.tr.setSelection(new AllSelection(state.doc))),
      );
      view.typeText("xyz");
      const accepted = paragraphs(resolveAllChanges(view.state, "accept").doc);
      expect(accepted).toEqual([
        {
          text: "xyz",
          paraId: "4C000004",
          alignment: "right",
          styleId: "BlockQuote",
          mark: null,
          propertyChange: false,
        },
      ]);
    }
  });

  test("a table pasted over the last paragraph goes in front of it, which stays", async () => {
    const { document, state } = await open(INTRO + paragraph("2A000002", "Alpha beta gamma."));
    const placed = placeSelection(state, "Alpha beta gamma.", "paragraph") ?? state;
    const view = new HeadlessEditorView(placed);
    const source = await open(table(), "editing");
    view.paste(source.state.doc.slice(0, source.state.doc.content.size));
    const kinds = (doc: import("prosemirror-model").Node) => {
      const names: string[] = [];
      doc.forEach((child) =>
        names.push(child.type.name === "paragraph" ? `p:${child.textContent}` : child.type.name),
      );
      return names;
    };
    expect(kinds(resolveAllChanges(view.state, "accept").doc)).toEqual([
      "p:Intro paragraph.",
      "table",
      "p:",
    ]);
    expect(kinds(resolveAllChanges(view.state, "reject").doc)).toEqual([
      "p:Intro paragraph.",
      "p:Alpha beta gamma.",
    ]);
    await saveHarnessState(view.state, document);
  });

  test("paragraphs pasted over the last paragraph record one property change on the last", async () => {
    const { document, state } = await open(INTRO + paragraph("2A000002", "Alpha beta gamma."));
    const placed = placeSelection(state, "Alpha beta gamma.", "paragraph") ?? state;
    const view = new HeadlessEditorView(placed);
    // The copied last paragraph already carries a pending property change.
    const pending =
      '<w:jc w:val="right"/><w:pPrChange w:id="9" w:author="Other" w:date="2026-01-01T00:00:00Z"><w:pPr><w:jc w:val="center"/></w:pPr></w:pPrChange>';
    const source = await open(
      paragraph("3B000003", "Bravo text.") + paragraph("4C000004", "Charlie text.", pending),
      "editing",
    );
    view.paste(source.state.doc.slice(0, source.state.doc.content.size));
    const last = textblocks(view.state.doc).at(-1)?.node;
    expect(last?.textContent).toBe("Charlie text.");
    expect(last ? (expectParagraphAttrs(last)._propertyChanges ?? []) : []).toHaveLength(1);
    expect(paragraphs(resolveAllChanges(view.state, "reject").doc)).toEqual(
      withoutIds(paragraphs(state.doc)),
    );
    expect(paragraphs(resolveAllChanges(view.state, "accept").doc).slice(1)).toMatchObject([
      { text: "Bravo text.", alignment: null },
      { text: "Charlie text.", alignment: "right", propertyChange: false },
    ]);
    await saveHarnessState(view.state, document);
  });

  test("undoing a table pasted over everything restores every paragraph", async () => {
    const { state } = await open(INTRO + CENTERED + TAIL);
    const view = new HeadlessEditorView(
      state.apply(state.tr.setSelection(new AllSelection(state.doc))),
    );
    const source = await open(table(), "editing");
    view.paste(source.state.doc.slice(0, source.state.doc.content.size));
    undo(view.state, view.dispatch);
    expect(paragraphs(view.state.doc)).toEqual(paragraphs(state.doc));
  });

  for (const key of ["Enter", "Shift-Enter"]) {
    test(`${key} at the end of a paragraph: rejecting leaves it as it was`, async () => {
      const { state } = await open(INTRO + CENTERED + TAIL);
      const view = new HeadlessEditorView(
        placeSelection(state, "Alpha text.", "caret-end") ?? state,
      );
      view.pressKey(key);
      expect(paragraphs(resolveAllChanges(view.state, "reject").doc)).toEqual(
        withoutIds(paragraphs(state.doc)),
      );
    });
  }

  test("paragraphs pasted at a caret: rejecting leaves the paragraph as it was", async () => {
    const { state } = await open(INTRO + CENTERED + TAIL);
    const view = new HeadlessEditorView(placeSelection(state, "Alpha text.", "caret-end") ?? state);
    const source = await open(
      RIGHT + paragraph("4C000004", "Charlie text.", '<w:jc w:val="right"/>'),
      "editing",
    );
    view.paste(source.state.doc.slice(1, source.state.doc.content.size - 1));
    expect(paragraphs(resolveAllChanges(view.state, "reject").doc)).toEqual(
      withoutIds(paragraphs(state.doc)),
    );
  });

  test("a paragraph whose mark goes runs on into the table after it", async () => {
    const deletedMark =
      '<w:p w14:paraId="2A000002" w14:textId="2A000002"><w:pPr><w:rPr><w:del w:id="9" w:author="Reviewer" w:date="2026-01-01T00:00:00Z"/></w:rPr></w:pPr><w:r><w:t>Alpha text.</w:t></w:r></w:p>';
    const { state } = await open(INTRO + deletedMark + table() + TAIL, "editing");
    const texts = (doc: import("prosemirror-model").Node) =>
      textblocks(doc).map(({ node }) => node.textContent);
    let single = state;
    acceptAIEditRevision(9)(state, (transaction) => {
      single = state.apply(transaction);
    });
    for (const accepted of [single, resolveAllChanges(state, "accept")]) {
      expect(texts(accepted.doc)).toEqual([
        "Intro paragraph.",
        "Alpha text.Cell.",
        "Tail paragraph.",
      ]);
      expect(paragraphs(accepted.doc)[1]).toMatchObject({ paraId: "66000001", mark: null });
    }
    expect(texts(resolveAllChanges(state, "reject").doc)).toEqual(texts(state.doc));
  });

  test("typing inside another author's insertion gives its second stretch an id of its own", async () => {
    const inserted =
      '<w:p w14:paraId="2A000002" w14:textId="2A000002"><w:ins w:id="7" w:author="Other" w:date="2026-01-01T00:00:00Z"><w:r><w:t>abcdef</w:t></w:r></w:ins></w:p>';
    const { state } = await open(INTRO + inserted);
    const target = textblocks(state.doc).find(({ node }) => node.textContent === "abcdef");
    if (!target) throw new Error("missing the inserted paragraph");
    const caret = target.pos + 1 + 3;
    const view = new HeadlessEditorView(
      state.apply(state.tr.setSelection(TextSelection.create(state.doc, caret))),
    );
    view.typeText("X");
    const ids: number[] = [];
    view.state.doc.nodesBetween(target.pos, target.pos + target.node.nodeSize + 1, (node) => {
      for (const mark of node.marks) {
        if (mark.type.name === "insertion") ids.push(mark.attrs["revisionId"] as number);
      }
    });
    expect(ids).toHaveLength(3);
    expect(new Set(ids).size).toBe(3);
  });
});

describe("runs a join moves read from the surviving paragraph's style", () => {
  const HEADING = paragraph("1A000001", "Service Agreement", '<w:pStyle w:val="Heading1"/>');
  const BODY = paragraph("2A000002", "This agreement is made between the parties named below.");
  const HEADING_RUN = { bold: true, size: 32 };
  const BODY_RUN = { bold: false, size: 22 };

  const caretIn = (view: HeadlessEditorView, index: number, at: "start" | "end" | number) => {
    const block = textblocks(view.state.doc)[index];
    if (!block) throw new Error(`missing paragraph ${index}`);
    let offset = at;
    if (at === "start") offset = 0;
    if (at === "end") offset = block.node.content.size;
    view.dispatch(
      view.state.tr.setSelection(
        TextSelection.create(view.state.doc, block.pos + 1 + Number(offset)),
      ),
    );
  };

  /** The saved body, and every run that states run properties of its own. */
  const savedBody = async (state: EditorState, base: Document) => {
    const saved = await saveHarnessState(state, base);
    const xml = await (
      await JSZip.loadAsync(saved.bytes)
    )
      .file("word/document.xml")
      ?.async("string");
    if (!xml) throw new Error("missing document.xml");
    const body = xml.slice(xml.indexOf("<w:body>"), xml.lastIndexOf("<w:sectPr"));
    return {
      bytes: saved.bytes,
      body,
      runProperties: body.match(/<w:r>(?:(?!<\/w:r>).)*?<w:rPr>.*?<\/w:rPr>/g) ?? [],
    };
  };

  /** Each reopened paragraph's style and how each of its runs reads. */
  const reopened = async (bytes: Uint8Array) => {
    const state = createHarnessState(await parseShapeDocument(bytes), "editing");
    return textblocks(state.doc).map(({ node }) => {
      const runs: { text: string; bold: boolean; size: unknown }[] = [];
      node.forEach((child) => {
        if (!child.isText) return;
        runs.push({
          text: child.text ?? "",
          bold: child.marks.some((mark) => mark.type.name === "bold"),
          size: child.marks.find((mark) => mark.type.name === "fontSize")?.attrs["size"],
        });
      });
      return { styleId: expectParagraphAttrs(node).styleId ?? null, runs };
    });
  };

  for (const mode of ["suggesting", "editing"] as const) {
    for (const lastJoin of ["Delete", "Backspace"] as const) {
      test(`${mode}: split, join into the heading, then ${lastJoin} the rest back on`, async () => {
        const { document, state } = await open(HEADING + BODY, mode);
        const view = new HeadlessEditorView(state);
        caretIn(view, 1, "This ".length);
        view.pressKey("Enter");
        caretIn(view, 1, "start");
        view.pressKey("Backspace");
        const part = textblocks(view.state.doc).findIndex(({ node }) =>
          node.textContent.endsWith("This "),
        );
        if (lastJoin === "Delete") {
          caretIn(view, part, "end");
          view.pressKey("Delete");
        } else {
          caretIn(view, part + 1, "start");
          view.pressKey("Backspace");
        }
        const merged = textblocks(view.state.doc).at(-1)?.node;
        expect(
          merged?.textContent.endsWith("This agreement is made between the parties named below."),
        ).toBe(true);
        // The paragraph left paints the heading's style-supplied properties.
        expect(merged ? expectParagraphAttrs(merged).keepNext : null).toBe(true);

        const saved = await savedBody(view.state, document);
        expect(saved.runProperties).toEqual([]);
        expect(saved.body).not.toContain("<w:keepNext");
        expect(saved.body).not.toContain("<w:outlineLvl");

        const accepted = await savedBody(resolveAllChanges(view.state, "accept"), document);
        expect(accepted.runProperties).toEqual([]);
        expect(accepted.body).not.toContain("<w:keepNext");
        const acceptedParagraphs = await reopened(accepted.bytes);
        expect(acceptedParagraphs).toHaveLength(1);
        expect(acceptedParagraphs[0]?.styleId).toBe("Heading1");
        expect(acceptedParagraphs[0]?.runs.map(({ text }) => text).join("")).toBe(
          "Service AgreementThis agreement is made between the parties named below.",
        );
        for (const run of acceptedParagraphs[0]?.runs ?? []) {
          expect(run).toMatchObject(HEADING_RUN);
        }

        if (mode === "suggesting") {
          const rejected = await savedBody(resolveAllChanges(view.state, "reject"), document);
          expect(rejected.runProperties).toEqual([]);
          expect(rejected.body).not.toContain("<w:keepNext");
          expect(await reopened(rejected.bytes)).toEqual([
            { styleId: "Heading1", runs: [{ text: "Service Agreement", ...HEADING_RUN }] },
            {
              styleId: null,
              runs: [
                { text: "This agreement is made between the parties named below.", ...BODY_RUN },
              ],
            },
          ]);
        }
      });
    }
  }

  test("a heading joined onto body text records and restores only its own properties", async () => {
    const { document, state } = await open(paragraph("3A000003", "Intro text.") + HEADING);
    const view = new HeadlessEditorView(state);
    caretIn(view, 1, "start");
    view.pressKey("Backspace");
    const saved = await savedBody(view.state, document);
    expect(saved.body).toContain("<w:pPrChange");
    expect(saved.body).not.toContain("<w:keepNext");
    expect(saved.body).not.toContain("<w:outlineLvl");
    expect(saved.runProperties).toEqual([]);

    const rejectedState = resolveAllChanges(view.state, "reject");
    const heading = textblocks(rejectedState.doc).at(-1)?.node;
    expect(heading ? expectParagraphAttrs(heading) : null).toMatchObject({
      styleId: "Heading1",
      keepNext: true,
    });
    const rejected = await savedBody(rejectedState, document);
    expect(rejected.body).not.toContain("<w:keepNext");
    expect(rejected.body).not.toContain("<w:outlineLvl");
    expect(rejected.runProperties).toEqual([]);
    expect(await reopened(rejected.bytes)).toEqual([
      { styleId: null, runs: [{ text: "Intro text.", ...BODY_RUN }] },
      { styleId: "Heading1", runs: [{ text: "Service Agreement", ...HEADING_RUN }] },
    ]);
  });
});
