/**
 * Footnote Reference Mark Extension
 *
 * Provides footnoteRef mark + insert/delete commands for footnotes and endnotes.
 */

import { panic } from "better-result";
import type { Command } from "prosemirror-state";

import { expectFootnoteRefMarkAttrs } from "../../attrs";
import { suggestRangeDeletion, suggestionModeKey } from "../../plugins/suggestionMode";
import { createMarkExtension } from "../create";
import type { ExtensionRuntime } from "../types";
import { expandNoteReferenceDeletionRange, noteReferenceRanges } from "./noteReferenceDeletion";

const noteRefAttrsFromDom = (
  dom: HTMLElement,
  noteType: "footnote" | "endnote",
  vertAlign?: "baseline" | "superscript",
): Record<string, string> => ({
  id: dom.dataset["id"] ?? "",
  noteType: dom.dataset["noteType"] ?? noteType,
  ...(vertAlign ? { vertAlign } : {}),
});

export const FootnoteRefExtension = createMarkExtension({
  name: "footnoteRef",
  schemaMarkName: "footnoteRef",
  markSpec: {
    inclusive: false,
    attrs: {
      id: {},
      noteType: { default: "footnote" },
      vertAlign: { default: null },
      customMarkFollows: { default: null },
    },
    parseDOM: [
      {
        tag: "sup.docx-footnote-ref",
        getAttrs: (dom) => noteRefAttrsFromDom(dom, "footnote", "superscript"),
      },
      {
        tag: "sup.docx-endnote-ref",
        getAttrs: (dom) => noteRefAttrsFromDom(dom, "endnote", "superscript"),
      },
      {
        tag: "span.docx-footnote-ref",
        getAttrs: (dom) => noteRefAttrsFromDom(dom, "footnote", "baseline"),
      },
      {
        tag: "span.docx-endnote-ref",
        getAttrs: (dom) => noteRefAttrsFromDom(dom, "endnote", "baseline"),
      },
    ],
    toDOM(mark) {
      const attrs = expectFootnoteRefMarkAttrs(mark);
      const id = String(attrs.id);
      const noteType = attrs.noteType ?? "footnote";
      const isSuperscript = attrs.vertAlign === "superscript";
      const tagName = isSuperscript ? "sup" : "span";
      const alignClass = isSuperscript ? "docx-note-ref-superscript" : "docx-note-ref-baseline";
      return [
        tagName,
        {
          class: `docx-${noteType}-ref ${alignClass}`,
          "data-id": id,
          "data-note-type": noteType,
        },
        0,
      ];
    },
  },
  onSchemaReady(): ExtensionRuntime {
    const deleteWholeNoteReference =
      (direction: "backward" | "forward"): Command =>
      (state, dispatch) => {
        if (suggestionModeKey.getState(state)?.active) {
          return false;
        }
        const { from, to, empty } = state.selection;
        const range = expandNoteReferenceDeletionRange(
          state.doc,
          empty && direction === "backward" ? Math.max(0, from - 1) : from,
          empty && direction === "forward" ? Math.min(state.doc.content.size, to + 1) : to,
        );
        if (!range) {
          return false;
        }
        dispatch?.(state.tr.delete(range.from, range.to).scrollIntoView());
        return true;
      };

    function makeInsertNote(noteType: "footnote" | "endnote"): (id: number) => Command {
      return (id: number): Command =>
        (state, dispatch) => {
          if (!dispatch) {
            return true;
          }

          const { schema } = state;
          const footnoteRefType = schema.marks["footnoteRef"];
          if (!footnoteRefType) {
            panic("Missing mark type: footnoteRef");
          }
          const mark = footnoteRefType.create({
            id: String(id),
            noteType,
            vertAlign: "superscript",
          });
          const text = schema.text(String(id), [mark]);
          const tr = state.tr.replaceSelectionWith(text, false);
          dispatch(tr.scrollIntoView());
          return true;
        };
    }

    /**
     * Delete the note references the selection touches (the one beside a
     * caret), each whole: tracked as a deletion of the reference run when
     * suggesting, removed outright otherwise.
     */
    const deleteNoteRef: Command = (state, dispatch) => {
      const { from, to, empty } = state.selection;
      const references = noteReferenceRanges(
        state.doc,
        empty ? Math.max(0, from - 1) : from,
        empty ? Math.min(state.doc.content.size, to + 1) : to,
      );
      if (references.length === 0) {
        return false;
      }
      if (!dispatch) {
        return true;
      }
      const tr = state.tr;
      for (const reference of references.toReversed()) {
        if (!suggestRangeDeletion(state, tr, reference.from, reference.to)) {
          tr.delete(reference.from, reference.to);
        }
      }
      dispatch(tr.scrollIntoView());
      return true;
    };

    return {
      keyboardShortcuts: {
        Backspace: deleteWholeNoteReference("backward"),
        Delete: deleteWholeNoteReference("forward"),
      },
      commands: {
        insertFootnote: makeInsertNote("footnote"),
        insertEndnote: makeInsertNote("endnote"),
        deleteNoteRef: () => deleteNoteRef,
      },
    };
  },
});
