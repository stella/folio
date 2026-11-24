import { history, undo } from "prosemirror-history";
import { keymap } from "prosemirror-keymap";
import { Node, Schema } from "prosemirror-model";
import { EditorState } from "prosemirror-state";
import { Transform } from "prosemirror-transform";
import { EditorView } from "prosemirror-view";
import * as yjs from "yjs";

import { schema as folioSchema } from "@stll/folio-core/prosemirror/schema";

const schema = new Schema({
  nodes: {
    doc: { content: "paragraph+" },
    paragraph: { content: "text*", toDOM: () => ["p", 0] },
    text: {},
  },
});

export const mountHostEditor = (element: HTMLElement) => {
  const view = new EditorView(element, {
    state: EditorState.create({
      schema,
      doc: schema.node("doc", null, [
        schema.node("paragraph", null, schema.text("Host text")),
      ]),
      plugins: [history(), keymap({ "Mod-z": undo })],
    }),
  });
  const doc = new yjs.Doc();
  view.dom.style.color = "rgb(12, 34, 56)";
  view.dom.style.caretColor = "rgb(12, 34, 56)";
  view.dom.style.outline = "3px solid rgb(12, 34, 56)";
  doc.getText("host").insert(0, "Host Yjs");
  return {
    view,
    identities(folioView: EditorView) {
      return {
        schema:
          folioSchema instanceof Schema &&
          folioView.state.schema instanceof Schema,
        node: folioView.state.doc instanceof Node,
        state: folioView.state instanceof EditorState,
        view: folioView instanceof EditorView,
        transform: folioView.state.tr instanceof Transform,
      };
    },
    text: () => view.state.doc.textContent,
    yjsText: () => doc.getText("host").toString(),
    destroy: () => {
      view.destroy();
      doc.destroy();
    },
  };
};
