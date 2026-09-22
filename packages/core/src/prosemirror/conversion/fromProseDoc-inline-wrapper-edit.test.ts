/**
 * An edit inside a transparent inline wrapper keeps the wrapper.
 *
 * This is the case the source-paragraph replay never covered: once a paragraph
 * is edited, its markup is rebuilt from the editor, and the wrapper only comes
 * back if the save leg reads the `inlineWrapper` mark the edited text still
 * carries. Both editing modes are exercised, because a tracked edit adds
 * revisions to the same span while authored wrapper ownership must survive.
 */

import { describe, expect, test } from "bun:test";
import { EditorState, TextSelection } from "prosemirror-state";

import type {
  Document,
  InlineWrapper,
  Paragraph,
  ParagraphContent,
  Run,
} from "../../types/document";
import { createEmptyDocument } from "../../utils/createDocument";
import { rejectAllChanges } from "../commands/comments";
import { createSuggestionModePlugin } from "../plugins/suggestionMode";
import { fromProseDoc } from "./fromProseDoc";
import { toProseDoc } from "./toProseDoc";

const run = (text: string): Run => ({ type: "run", content: [{ type: "text", text }] });

const WRAPPED: InlineWrapper = {
  type: "inlineWrapper",
  kind: "bidi",
  control: "override",
  direction: "rtl",
  content: [run("inside")],
};

/** The same shape for the kinds that name an element rather than a layout. */
const TAGGED: readonly InlineWrapper[] = [
  {
    type: "inlineWrapper",
    kind: "smartTag",
    element: "City",
    uri: "urn:example:tags",
    propertiesXml: '<w:smartTagPr><w:attr w:name="k" w:val="v"/></w:smartTagPr>',
    content: [run("inside")],
  },
  {
    type: "inlineWrapper",
    kind: "customXml",
    element: "party",
    content: [run("inside")],
  },
];

const REVISION_INFO = { id: 7, author: "Reviewer", date: "2026-01-01T00:00:00Z" };

/** `w:bdo > w:ins > w:r`: the wrapper the author put outside the revision. */
const WRAPPER_OUTSIDE_REVISION: InlineWrapper = {
  type: "inlineWrapper",
  kind: "bidi",
  control: "override",
  direction: "rtl",
  content: [{ type: "insertion", info: REVISION_INFO, content: [run("inside")] }],
};

/** `w:ins > w:bdo > w:r`: the same span with the revision outside. */
const REVISION_OUTSIDE_WRAPPER: ParagraphContent = {
  type: "insertion",
  info: REVISION_INFO,
  content: [WRAPPED],
};

const documentWith = (content: Paragraph["content"]): Document => {
  const template = createEmptyDocument();
  return {
    ...template,
    package: {
      ...template.package,
      document: {
        ...template.package.document,
        content: [{ type: "paragraph", paraId: "C0000001", content }],
      },
    },
  };
};

const paragraphContentOf = (document: Document): ParagraphContent[] => {
  const block = document.package.document.content.at(0);
  if (block?.type !== "paragraph") {
    throw new Error("The rebuilt document lost its paragraph");
  }
  return block.content;
};

const rangeOfText = (state: EditorState, text: string): { from: number; to: number } => {
  let range: { from: number; to: number } | undefined;
  state.doc.descendants((node, position) => {
    const at = node.isText ? (node.text?.indexOf(text) ?? -1) : -1;
    if (range === undefined && at >= 0) {
      range = { from: position + at, to: position + at + text.length };
    }
  });
  if (range === undefined) {
    throw new Error(`The editor holds no text ${text}`);
  }
  return range;
};

type EditMode = "direct" | "tracked";

/** Replace `target` with `replacement` inside the editor, then save. */
const editAndSave = (
  content: Paragraph["content"],
  mode: EditMode,
  target: string,
  replacement: string,
): ParagraphContent[] => {
  const source = documentWith(content);
  const pmDoc = toProseDoc(source);
  const state = EditorState.create({
    doc: pmDoc,
    plugins: [createSuggestionModePlugin(mode === "tracked", "Reviewer")],
  });
  const { from, to } = rangeOfText(state, target);
  const selected = state.apply(state.tr.setSelection(TextSelection.create(state.doc, from, to)));
  const edited = selected.apply(selected.tr.insertText(replacement, from, to));
  return paragraphContentOf(fromProseDoc(edited.doc, source));
};

/** Every wrapper in `content`, at any depth, with what it holds. */
const wrappersIn = (content: readonly ParagraphContent[]): InlineWrapper[] => {
  const found: InlineWrapper[] = [];
  const visit = (items: readonly ParagraphContent[]): void => {
    for (const item of items) {
      if (item.type === "inlineWrapper") {
        found.push(item);
        visit(item.content);
        continue;
      }
      if (
        item.type === "insertion" ||
        item.type === "deletion" ||
        item.type === "moveFrom" ||
        item.type === "moveTo"
      ) {
        visit(item.content);
      }
    }
  };
  visit(content);
  return found;
};

const textIn = (content: readonly ParagraphContent[]): string => {
  let text = "";
  const visit = (items: readonly ParagraphContent[]): void => {
    for (const item of items) {
      switch (item.type) {
        case "run":
          for (const part of item.content) {
            if (part.type === "text") {
              text += part.text;
            }
          }
          break;
        case "inlineWrapper":
        case "insertion":
        case "deletion":
        case "moveFrom":
        case "moveTo":
          visit(item.content);
          break;
        default:
          break;
      }
    }
  };
  visit(content);
  return text;
};

describe("text replaced inside a bidirectional wrapper", () => {
  for (const mode of ["direct", "tracked"] as const) {
    test(`keeps the wrapper (${mode})`, () => {
      const saved = editAndSave([WRAPPED, run(" outside")], mode, "sid", "SID");
      const wrappers = wrappersIn(saved);
      // A tracked edit cuts the span into a kept part, a deletion and an
      // insertion. What the wrapper says is the same in every one of them.
      expect(wrappers.length).toBeGreaterThan(0);
      for (const wrapper of wrappers) {
        expect(wrapper.control).toBe("override");
        expect(wrapper.direction).toBe("rtl");
      }
      expect(textIn(wrappers)).toContain("SID");
      // The text outside the wrapper stays outside it.
      expect(textIn(wrappers)).not.toContain("outside");
    });

    test(`keeps the authored wrapper and revision order (${mode})`, () => {
      for (const authored of [WRAPPER_OUTSIDE_REVISION, REVISION_OUTSIDE_WRAPPER]) {
        const saved = editAndSave([authored], mode, "sid", "SID");
        expect(saved.at(0)?.type).toBe(authored.type);
        expect(wrappersIn(saved).length).toBeGreaterThan(0);
      }
    });

    test(`saving the result again does not change it (${mode})`, () => {
      const once = editAndSave([WRAPPED, run(" outside")], mode, "sid", "SID");
      const source = documentWith(once);
      const twice = paragraphContentOf(fromProseDoc(toProseDoc(source), source));
      expect(twice).toEqual(once);
    });
  }
});

describe("text replaced inside a smart tag or a custom-XML wrapper", () => {
  for (const authored of TAGGED) {
    if (authored.kind === "bidi") {
      throw new Error("The tagged fixtures are not bidirectional wrappers");
    }
    for (const mode of ["direct", "tracked"] as const) {
      test(`keeps the ${authored.kind} around the replacement (${mode})`, () => {
        const wrappers = wrappersIn(editAndSave([authored, run(" outside")], mode, "sid", "SID"));
        expect(wrappers.length).toBeGreaterThan(0);
        for (const wrapper of wrappers) {
          if (wrapper.kind === "bidi") {
            throw new Error("The edit turned the wrapper into a bidirectional one");
          }
          expect(wrapper.kind).toBe(authored.kind);
          expect(wrapper.element).toBe(authored.element);
          expect(wrapper.uri).toBe(authored.uri);
          // The properties belong to the tag, so every part the tracked edit
          // cut the span into carries the same ones.
          expect(wrapper.propertiesXml).toBe(authored.propertiesXml);
        }
        expect(textIn(wrappers)).toContain("SID");
        expect(textIn(wrappers)).not.toContain("outside");
      });

      test(`saving the ${authored.kind} result again does not change it (${mode})`, () => {
        const once = editAndSave([authored, run(" outside")], mode, "sid", "SID");
        const source = documentWith(once);
        expect(paragraphContentOf(fromProseDoc(toProseDoc(source), source))).toEqual(once);
      });
    }
  }
});

describe("a wrapper with nothing left in it", () => {
  test("is not written once its text is deleted", () => {
    const source = documentWith([WRAPPED, run(" outside")]);
    const state = EditorState.create({
      doc: toProseDoc(source),
      plugins: [createSuggestionModePlugin(false, "Reviewer")],
    });
    const { from, to } = rangeOfText(state, "inside");
    const saved = paragraphContentOf(
      fromProseDoc(state.apply(state.tr.delete(from, to)).doc, source),
    );
    expect(wrappersIn(saved)).toHaveLength(0);
    expect(textIn(saved)).toBe(" outside");
  });

  test("is not written once the insertion that held it is rejected", () => {
    const source = documentWith([REVISION_OUTSIDE_WRAPPER, run(" outside")]);
    const state = EditorState.create({ doc: toProseDoc(source) });
    let rejected = state;
    rejectAllChanges()(state, (tr) => {
      rejected = state.apply(tr);
    });
    const saved = paragraphContentOf(fromProseDoc(rejected.doc, source));
    expect(wrappersIn(saved)).toHaveLength(0);
    expect(textIn(saved)).toBe(" outside");
  });
});
