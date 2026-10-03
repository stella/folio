/** An authored automatic note mark: invisible, but retained through PM JSON and edits. */
import { expectNoteMarkerAttrs } from "../../../internal/noteMarkerAttrs";
import { createNodeExtension } from "../create";

export const NOTE_MARKER_NODE_NAME = "noteMarker";

export const NoteMarkerExtension = createNodeExtension({
  name: NOTE_MARKER_NODE_NAME,
  schemaNodeName: NOTE_MARKER_NODE_NAME,
  nodeSpec: {
    inline: true,
    group: "inline",
    atom: true,
    marks: "_",
    selectable: false,
    attrs: { kind: {} },
    parseDOM: [
      {
        tag: "span[data-docx-note-marker]",
        getAttrs(node) {
          if (!(node instanceof HTMLElement)) return false;
          const kind = node.dataset["docxNoteMarker"];
          return kind === "footnote" || kind === "endnote" ? { kind } : false;
        },
      },
    ],
    toDOM(node) {
      const { kind } = expectNoteMarkerAttrs(node);
      return [
        "span",
        {
          "data-docx-note-marker": kind,
          "aria-hidden": "true",
          contenteditable: "false",
          style: "display: none;",
        },
      ];
    },
  },
});
