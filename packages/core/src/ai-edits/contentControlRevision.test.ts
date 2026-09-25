/**
 * An edit that removes all of a content control's text revises the text, not
 * the control.
 *
 * `w:sdt > w:sdtContent > w:del` deletes the control's content; accepting it
 * leaves the control standing, emptied, where a consumer shows its placeholder
 * (ECMA-376 §17.5.2, `w:placeholder`). Only a revision around the control
 * (`w:del > w:sdt`), a deletion spanning the whole control in the editor, or a
 * deleted paragraph removes it.
 */

import { describe, expect, test } from "bun:test";
import JSZip from "jszip";
import type { Node as PMNode } from "prosemirror-model";
import { EditorState, TextSelection, type Transaction } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";

import { RELATIONSHIP_TYPES } from "../docx/relsParser";
import { acceptAllChanges, rejectAllChanges } from "../prosemirror/commands/comments";
import { createSuggestionModePlugin } from "../prosemirror/plugins/suggestionMode";
import { schema } from "../prosemirror/schema";
import { FolioDocxReviewer } from "./headless";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

const createDocx = async (bodyXml: string): Promise<ArrayBuffer> => {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    `${XML_DECLARATION}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
      `<Default Extension="xml" ContentType="application/xml"/>` +
      `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
      `</Types>`,
  );
  zip.file(
    "_rels/.rels",
    `${XML_DECLARATION}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="${RELATIONSHIP_TYPES.officeDocument}" Target="word/document.xml"/></Relationships>`,
  );
  zip.file(
    "word/document.xml",
    `${XML_DECLARATION}<w:document xmlns:w="${W}"><w:body>${bodyXml}` +
      `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`,
  );
  return zip.generateAsync({ type: "arraybuffer" });
};

const bodyOf = async (saved: ArrayBuffer): Promise<string> => {
  const xml = (await (await JSZip.loadAsync(saved)).file("word/document.xml")?.async("text")) ?? "";
  return xml.slice(xml.indexOf("<w:body>") + "<w:body>".length, xml.indexOf("<w:sectPr>"));
};

const CONTROL_PROPERTIES =
  '<w:sdtPr><w:alias w:val="Name"/><w:id w:val="7"/>' +
  '<w:placeholder><w:docPart w:val="DefaultPlaceholder"/></w:placeholder></w:sdtPr>';

const FIXTURES = [
  {
    kind: "inline",
    find: "John",
    bodyXml:
      '<w:p><w:r><w:t xml:space="preserve">Name: </w:t></w:r>' +
      `<w:sdt>${CONTROL_PROPERTIES}<w:sdtContent><w:r><w:t>John</w:t></w:r></w:sdtContent></w:sdt>` +
      '<w:r><w:t xml:space="preserve"> pays.</w:t></w:r></w:p>',
    acceptedText: "Name:  pays.",
  },
  {
    kind: "block",
    find: "Clause text",
    bodyXml:
      "<w:p><w:r><w:t>Before.</w:t></w:r></w:p>" +
      `<w:sdt>${CONTROL_PROPERTIES}<w:sdtContent>` +
      "<w:p><w:r><w:t>Clause text</w:t></w:r></w:p></w:sdtContent></w:sdt>" +
      "<w:p><w:r><w:t>After.</w:t></w:r></w:p>",
    acceptedText: "",
  },
] as const;

const MODES = ["tracked-changes", "suggested"] as const;

const deleteControlText = async (
  fixture: (typeof FIXTURES)[number],
  mode: (typeof MODES)[number],
) => {
  const source = await createDocx(fixture.bodyXml);
  const original = await bodyOf(await (await FolioDocxReviewer.fromBuffer(source)).toBuffer());
  const reviewer = await FolioDocxReviewer.fromBuffer(source);
  const blockWithControl = reviewer
    .snapshot()
    .blocks.find((candidate) => candidate.text.includes(fixture.find));
  if (!blockWithControl) {
    throw new Error("expected the control's block");
  }
  const result = reviewer.applyOperations(
    [
      {
        id: "empty",
        type: "replaceInBlock",
        blockId: blockWithControl.id,
        find: fixture.find,
        replace: "",
      },
    ],
    { mode },
  );
  expect(result.skipped).toEqual([]);
  return { reviewer, original, blockId: blockWithControl.id };
};

const controlCount = (xml: string): number => (xml.match(/<w:sdt>/gu) ?? []).length;

describe("deleting all of a content control's text", () => {
  for (const fixture of FIXTURES) {
    for (const mode of MODES) {
      test(`accepting leaves the ${fixture.kind} control, emptied (${mode})`, async () => {
        const { reviewer, blockId } = await deleteControlText(fixture, mode);
        reviewer.acceptAll();
        const accepted = await bodyOf(await reviewer.toBuffer());
        expect(controlCount(accepted)).toBe(1);
        expect(accepted).toContain('<w:alias w:val="Name"/>');
        expect(accepted).toContain('<w:docPart w:val="DefaultPlaceholder"/>');
        expect(accepted).not.toContain(fixture.find);
        expect(reviewer.snapshot().blocks.find((block) => block.id === blockId)?.text).toBe(
          fixture.acceptedText,
        );
      });

      test(`rejecting restores the ${fixture.kind} control (${mode})`, async () => {
        const { reviewer, original } = await deleteControlText(fixture, mode);
        reviewer.rejectAll();
        expect(await bodyOf(await reviewer.toBuffer())).toBe(original);
      });
    }

    test(`the saved redline keeps the ${fixture.kind} deletion inside the control`, async () => {
      const { reviewer } = await deleteControlText(fixture, "tracked-changes");
      const saved = await reviewer.toBuffer();
      expect(await bodyOf(saved)).toMatch(/<w:sdtContent>(?:<w:p>)?<w:del /u);
      const reopened = await FolioDocxReviewer.fromBuffer(saved);
      reopened.acceptAll();
      expect(controlCount(await bodyOf(await reopened.toBuffer()))).toBe(1);
    });
  }

  test("accepting one change by its id leaves the control too", async () => {
    const { reviewer } = await deleteControlText(FIXTURES[0], "tracked-changes");
    for (const change of reviewer.getChanges()) {
      reviewer.acceptChange(change);
    }
    const accepted = await bodyOf(await reviewer.toBuffer());
    expect(controlCount(accepted)).toBe(1);
    expect(accepted).not.toContain("John");
  });

  test("a deleted paragraph takes its controls with it", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(
      await createDocx(FIXTURES[0].bodyXml + "<w:p><w:r><w:t>Next.</w:t></w:r></w:p>"),
    );
    const block = reviewer.snapshot().blocks.at(0);
    if (!block) {
      throw new Error("expected a block");
    }
    reviewer.applyOperations([{ id: "drop", type: "deleteBlock", blockId: block.id }], {
      mode: "tracked-changes",
    });
    const saved = await reviewer.toBuffer();
    // The paragraph's deletion is written around the control, not inside it.
    expect(await bodyOf(saved)).toContain("<w:sdtContent><w:r><w:delText>John</w:delText>");
    for (const accepting of [reviewer, await FolioDocxReviewer.fromBuffer(saved)]) {
      accepting.acceptAll();
      const accepted = await bodyOf(await accepting.toBuffer());
      expect(controlCount(accepted)).toBe(0);
      expect(accepting.snapshot().blocks.map((candidate) => candidate.text)).toEqual(["Next."]);
    }
  });
});

describe("a tracked deletion typed in the editor", () => {
  const CONTROL_TEXT = "John";

  /** `Name: [John] pays.` with suggestion mode on; returns the control's position. */
  const editorState = () => {
    const control = schema.node("sdt", { sdtType: "richText", tag: "bound" }, [
      schema.text(CONTROL_TEXT),
    ]);
    const paragraph = schema.node("paragraph", null, [
      schema.text("Name: "),
      control,
      schema.text(" pays."),
    ]);
    const state = EditorState.create({
      schema,
      doc: schema.node("doc", null, [paragraph]),
      plugins: [createSuggestionModePlugin(true, "Reviewer")],
    });
    const controlPos = 1 + "Name: ".length;
    return { state, controlPos, controlEnd: controlPos + control.nodeSize };
  };

  const backspaceOver = (state: EditorState, from: number, to: number): EditorState => {
    let next = state.apply(state.tr.setSelection(TextSelection.create(state.doc, from, to)));
    const view = {
      get state() {
        return next;
      },
      dispatch(transaction: Transaction) {
        next = next.apply(transaction);
      },
    };
    // SAFETY: the suggestion-mode key handler reads only the view's `state`
    // and `dispatch`, and only the event's `key`.
    const editorView = view as unknown as EditorView;
    const event = { key: "Backspace" } as KeyboardEvent;
    const handled = next.plugins
      .map((plugin) => plugin.props.handleKeyDown?.call(plugin, editorView, event))
      .some(Boolean);
    expect(handled).toBe(true);
    return next;
  };

  const resolve = (state: EditorState, command: typeof acceptAllChanges): PMNode => {
    let next = state;
    command()(state, (transaction) => {
      next = state.apply(transaction);
    });
    return next.doc;
  };

  const controlsIn = (doc: PMNode): string[] => {
    const controls: string[] = [];
    doc.descendants((node) => {
      if (node.type.name === "sdt") {
        controls.push(node.textContent);
      }
      return true;
    });
    return controls;
  };

  test("over the control's text only, accepting leaves the control, emptied", () => {
    const { state, controlPos } = editorState();
    const deleted = backspaceOver(state, controlPos + 1, controlPos + 1 + CONTROL_TEXT.length);
    expect(controlsIn(resolve(deleted, acceptAllChanges))).toEqual([""]);
    expect(resolve(deleted, rejectAllChanges).eq(state.doc)).toBe(true);
  });

  test("across the whole control, accepting removes it as deleting it directly would", () => {
    const { state, controlPos, controlEnd } = editorState();
    const deleted = backspaceOver(state, controlPos - 1, controlEnd + 1);
    const accepted = resolve(deleted, acceptAllChanges);
    expect(controlsIn(accepted)).toEqual([]);
    expect(accepted.textContent).toBe("Name:pays.");
    expect(controlsIn(resolve(deleted, rejectAllChanges))).toEqual([CONTROL_TEXT]);
  });
});
