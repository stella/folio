import { describe, expect, test } from "bun:test";
import { EditorState, TextSelection } from "prosemirror-state";
import { MAX_REVISION_ID } from "@stll/docx-core/model";
import { panic } from "better-result";

import { createNumberingMap } from "../docx/numberingParser";
import { fromMarkdown } from "../markdown/fromMarkdown";
import type { NumberingDefinitions } from "../types/document";
import { toProseDoc } from "./conversion/toProseDoc";
import { toggleBulletList, toggleNumberedList } from "./extensions/features/ListExtension";
import { completeNumberingForDoc } from "./listInstanceReferences";
import {
  paragraphNumberingReference,
  paragraphNumberingReferenceId,
} from "../docx/numberingReference";
import { paragraphNumberingAttr } from "./numberingAttr";
import { expectParagraphAttrs } from "./attrs";
import { createDocumentNumberingPlugin, getDocumentNumbering } from "./plugins/documentNumbering";
import { storyListNumbering } from "./storyListNumbering";

/** A one-paragraph story editor over a package that defines `definitions`. */
const storyState = (definitions: NumberingDefinitions | undefined): EditorState => {
  const state = EditorState.create({
    doc: toProseDoc(fromMarkdown("Story line")),
    plugins: [createDocumentNumberingPlugin(definitions)],
  });
  return state.apply(state.tr.setSelection(TextSelection.create(state.doc, 1)));
};

const toggled = (state: EditorState, command: typeof toggleBulletList): EditorState => {
  let next = state;
  command(state, (tr) => {
    next = state.apply(tr);
  });
  return next;
};

const firstNumId = (state: EditorState | { doc: EditorState["doc"] }): number | undefined => {
  const node = state.doc.firstChild;
  return node ? paragraphNumberingReferenceId(expectParagraphAttrs(node).numPr) : undefined;
};

describe("lists started in two stories", () => {
  test("a story's list moves to its own id when another story defined the same id", () => {
    // Independent stories allocate against the same package definitions.
    const body = toggled(storyState(undefined), toggleBulletList);
    const originalHeader = toggled(storyState(undefined), toggleNumberedList);
    const bodyNumId = firstNumId(body);
    if (bodyNumId === undefined) panic("Body list has no numbering id");
    const header = originalHeader.apply(
      originalHeader.tr.setNodeAttribute(
        0,
        "numPr",
        paragraphNumberingAttr(paragraphNumberingReference({ numId: bodyNumId, ilvl: 0 })),
      ),
    );
    expect(firstNumId(header)).toBe(firstNumId(body));

    const withBody = completeNumberingForDoc(undefined, body.doc);
    const story = storyListNumbering(header, withBody);
    const repeated = storyListNumbering(header, withBody);
    expect(repeated.doc.toJSON()).toEqual(story.doc.toJSON());
    expect(repeated.numbering).toEqual(story.numbering);

    const headerNumId = firstNumId(story);
    expect(headerNumId).not.toBe(firstNumId(body));
    const map = createNumberingMap(story.numbering ?? { abstractNums: [], nums: [] });
    expect(map.getLevel(firstNumId(body) ?? -1, 0)?.numFmt).toBe("bullet");
    expect(map.getLevel(headerNumId ?? -1, 0)?.numFmt).toBe("decimal");
  });

  test("a story saved again keeps its list", () => {
    const header = toggled(storyState(undefined), toggleNumberedList);
    const first = storyListNumbering(header, undefined);
    const second = storyListNumbering(header, first.numbering);

    expect(second.doc).toBe(header.doc);
    expect(firstNumId(second)).toBe(firstNumId(header));
  });

  test("colliding story lists wrap around loaded maximum ids without reusing definitions", () => {
    const body = toggled(storyState(undefined), toggleBulletList);
    const header = toggled(storyState(undefined), toggleNumberedList);
    const headerId = firstNumId(header);
    const bodyNumbering = completeNumberingForDoc(undefined, body.doc);
    const abstract = bodyNumbering?.abstractNums.at(0);
    const instance = bodyNumbering?.nums.at(0);
    if (headerId === undefined || abstract === undefined || instance === undefined) {
      panic("Story lists have no numbering definitions");
    }
    const numbering = {
      abstractNums: [abstract, { ...abstract, abstractNumId: MAX_REVISION_ID }],
      nums: [
        { ...instance, numId: headerId },
        { ...instance, numId: MAX_REVISION_ID },
      ],
    } satisfies NumberingDefinitions;
    const story = storyListNumbering(header, numbering);
    const remappedId = firstNumId(story);
    expect(remappedId).not.toBe(headerId);
    expect(remappedId).not.toBe(MAX_REVISION_ID);
    for (const id of [
      ...(story.numbering?.nums ?? []).map(({ numId }) => numId),
      ...(story.numbering?.abstractNums ?? []).map(({ abstractNumId }) => abstractNumId),
    ]) {
      expect(Number.isInteger(id)).toBe(true);
      expect(id).toBeGreaterThanOrEqual(0);
      expect(id).toBeLessThanOrEqual(MAX_REVISION_ID);
    }
  });
});

describe("document numbering state", () => {
  test("follows a list another client inserts, and ignores typing", () => {
    const state = storyState(undefined);
    const withList = toggled(storyState(undefined), toggleNumberedList);
    const listParagraph = withList.doc.firstChild;
    if (!listParagraph) panic("The toggled story has no paragraph");

    const typed = state.apply(state.tr.insertText("x", 1));
    expect(getDocumentNumbering(typed)).toBe(getDocumentNumbering(state));

    const inserted = typed.apply(typed.tr.insert(typed.doc.content.size, listParagraph));
    const numId = firstNumId(withList) ?? -1;
    expect(getDocumentNumbering(inserted)?.getLevel(numId, 0)?.numFmt).toBe("decimal");

    const attrState = storyState(undefined);
    const byAttr = attrState.apply(
      attrState.tr
        .setNodeAttribute(0, "numPr", listParagraph.attrs["numPr"])
        .setNodeAttribute(0, "listNumFmt", "decimal"),
    );
    expect(getDocumentNumbering(byAttr)?.hasNumbering(numId)).toBe(true);
  });
});
