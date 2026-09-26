import { describe, expect, test } from "bun:test";
import { EditorState, TextSelection } from "prosemirror-state";

import { createNumberingMap } from "../docx/numberingParser";
import { fromMarkdown } from "../markdown/fromMarkdown";
import type { NumberingDefinitions } from "../types/document";
import { toProseDoc } from "./conversion/toProseDoc";
import { toggleBulletList, toggleNumberedList } from "./extensions/features/ListExtension";
import { completeNumberingForDoc } from "./listInstanceReferences";
import { paragraphNumberingReferenceId } from "../docx/numberingReference";
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
    // Both stories start from a package without numbering and pick the same id.
    const body = toggled(storyState(undefined), toggleBulletList);
    const header = toggled(storyState(undefined), toggleNumberedList);
    expect(firstNumId(header)).toBe(firstNumId(body));

    const withBody = completeNumberingForDoc(undefined, body.doc);
    const story = storyListNumbering(header, withBody);

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
});

describe("document numbering state", () => {
  test("follows a list another client inserts, and ignores typing", () => {
    const state = storyState(undefined);
    const withList = toggled(storyState(undefined), toggleNumberedList);
    const listParagraph = withList.doc.firstChild;
    if (!listParagraph) throw new Error("the toggled story has no paragraph");

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
