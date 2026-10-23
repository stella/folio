/**
 * Footnote Reference Mark Extension
 *
 * Provides footnoteRef mark + insert/delete commands for footnotes and endnotes.
 */

import { panic } from "better-result";
import { type Command, TextSelection } from "prosemirror-state";

import { expectFootnoteRefMarkAttrs } from "../../attrs";
import { withCanonicalCommand } from "../../canonicalCommands";
import {
  mintNoteReferenceOccurrenceId,
  noteReferenceOccurrencePlugin,
} from "../../noteReferenceOccurrences";
import {
  suggestionModeKey,
  suggestRangeDeletion,
} from "../../plugins/suggestionMode";
import { createMarkExtension } from "../create";
import type { ExtensionRuntime } from "../types";
import {
  expandNoteReferenceDeletionRange,
  noteReferenceRanges,
} from "./noteReferenceDeletion";

const NOTE_REFERENCE_PARSE_PRIORITY = 60;

const noteRefAttrsFromDom = (
  dom: HTMLElement,
  noteType: "footnote" | "endnote",
  vertAlign?: "baseline" | "superscript",
): Record<string, string> => ({
  occurrenceId:
    dom.dataset["noteOccurrence"] ?? mintNoteReferenceOccurrenceId(),
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
      occurrenceId: {},
      noteType: { default: "footnote" },
      vertAlign: { default: null },
      customMarkFollows: { default: null },
    },
    parseDOM: [
      {
        tag: "sup.docx-footnote-ref",
        priority: NOTE_REFERENCE_PARSE_PRIORITY,
        getAttrs: (dom) => noteRefAttrsFromDom(dom, "footnote", "superscript"),
      },
      {
        tag: "sup.docx-endnote-ref",
        priority: NOTE_REFERENCE_PARSE_PRIORITY,
        getAttrs: (dom) => noteRefAttrsFromDom(dom, "endnote", "superscript"),
      },
      {
        tag: "span.docx-footnote-ref",
        priority: NOTE_REFERENCE_PARSE_PRIORITY,
        getAttrs: (dom) => noteRefAttrsFromDom(dom, "footnote", "baseline"),
      },
      {
        tag: "span.docx-endnote-ref",
        priority: NOTE_REFERENCE_PARSE_PRIORITY,
        getAttrs: (dom) => noteRefAttrsFromDom(dom, "endnote", "baseline"),
      },
    ],
    toDOM(mark) {
      const attrs = expectFootnoteRefMarkAttrs(mark);
      const id = String(attrs.id);
      const noteType = attrs.noteType ?? "footnote";
      const isSuperscript = attrs.vertAlign === "superscript";
      const tagName = isSuperscript ? "sup" : "span";
      const alignClass = isSuperscript
        ? "docx-note-ref-superscript"
        : "docx-note-ref-baseline";
      return [
        tagName,
        {
          class: `docx-${noteType}-ref ${alignClass}`,
          "data-id": id,
          "data-note-occurrence": attrs.occurrenceId,
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
          empty && direction === "forward"
            ? Math.min(state.doc.content.size, to + 1)
            : to,
        );
        if (!range) {
          return false;
        }
        dispatch?.(state.tr.delete(range.from, range.to).scrollIntoView());
        return true;
      };

    function makeInsertNote(
      noteType: "footnote" | "endnote",
    ): (id: number) => Command {
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
            occurrenceId: mintNoteReferenceOccurrenceId(),
            noteType,
            vertAlign: "superscript",
          });
          const text = schema.text(String(id), [mark]);
          const insertionEnd = state.selection.to + text.nodeSize;
          const tr = state.tr.insert(state.selection.to, text);
          tr.setSelection(TextSelection.create(tr.doc, insertionEnd));
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
      plugins: [noteReferenceOccurrencePlugin()],
      keyboardShortcuts: {
        Backspace: deleteWholeNoteReference("backward"),
        Delete: deleteWholeNoteReference("forward"),
      },
      commands: {
        insertFootnote: (id, content = [{ type: "paragraph", content: [] }]) =>
          withCanonicalCommand(makeInsertNote("footnote")(id), (state) => [
            {
              type: "insertNote",
              at: state.selection.to,
              kind: "footnote",
              id,
              content,
            },
          ]),
        insertEndnote: (id, content = [{ type: "paragraph", content: [] }]) =>
          withCanonicalCommand(makeInsertNote("endnote")(id), (state) => [
            {
              type: "insertNote",
              at: state.selection.to,
              kind: "endnote",
              id,
              content,
            },
          ]),
        deleteNoteRef: () =>
          withCanonicalCommand(deleteNoteRef, (state) => {
            const { from, to, empty } = state.selection;
            return noteReferenceRanges(
              state.doc,
              empty ? Math.max(0, from - 1) : from,
              empty ? Math.min(state.doc.content.size, to + 1) : to,
            ).flatMap((range) => {
              const node = state.doc.nodeAt(range.from);
              const mark = node?.marks.find(
                (candidate) => candidate.type.name === "footnoteRef",
              );
              if (!mark) return [];
              const attrs = expectFootnoteRefMarkAttrs(mark);
              const id = Number(attrs.id);
              if (!Number.isSafeInteger(id) || id <= 0) return [];
              return [
                {
                  type: "deleteNote" as const,
                  from: range.from,
                  kind: attrs.noteType ?? "footnote",
                  id,
                },
              ];
            });
          }),
      },
    };
  },
});
