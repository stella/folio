/**
 * A zero-width block for markup folio keeps opaquely between two blocks.
 *
 * The block-level twin of the `preservedXml` inline atom. A `w:permStart`
 * between two paragraphs, an `m:oMathPara`, a `w:altChunk`, a comment or move
 * range that opens between blocks: the editor has no model for any of it and
 * does not need one, because it only has to not lose it. So the node renders
 * nothing, takes no selection and holds no text.
 *
 * **Position is structure, not an attribute.** The node sits between the same
 * two blocks it sat between in the source, and ProseMirror's own mapping keeps
 * it there: inserting, splitting or deleting a neighbour moves the markup the
 * way a reader would expect, and nothing has to keep an index honest. The
 * alternative — an index into the container recorded on the node — drifts the
 * first time anything around it changes, and a drifting index is how markup
 * ends up in the wrong place rather than merely lost.
 *
 * Deleting a *neighbour* therefore leaves the markup attached to whatever now
 * follows it; only a selection that covers the node itself removes it, which
 * is the same rule the editor applies to every other block.
 *
 * The integrity plugin removes a node that can never write markup back: bad
 * attributes, or an empty payload left by a paste through the DOM. Keeping one
 * would leave a zero-width block in the document that serializes to nothing.
 */

import type { Node as PMNode } from "prosemirror-model";
import { Plugin } from "prosemirror-state";

import { expectPreservedBlockAttrs, readPreservedBlockAttrs } from "../../attrs";
import { createNodeExtension } from "../create";

/** The node name, for callers asking what a container holds. */
export const PRESERVED_BLOCK_NODE_NAME = "preservedBlock";

type PositionedNode = {
  position: number;
  node: PMNode;
};

/**
 * A node that can never write markup back.
 *
 * Either the attributes do not read at all, or `xml` is empty: a copy that
 * came back from a paste with its payload stripped. Keeping one would put a
 * zero-width block in the document that serializes to nothing, so the plugin
 * removes it rather than letting it ride along.
 */
const isUnwritable = (node: PMNode): boolean => {
  const attrs = readPreservedBlockAttrs(node);
  return !attrs.ok || attrs.value.xml === "";
};

const collectUnwritableBlocks = (doc: PMNode): PositionedNode[] => {
  const unwritable: PositionedNode[] = [];
  doc.descendants((node, position) => {
    if (node.type.name !== PRESERVED_BLOCK_NODE_NAME) {
      return true;
    }
    if (isUnwritable(node)) {
      unwritable.push({ position, node });
    }
    return false;
  });
  return unwritable;
};

export const PreservedBlockExtension = createNodeExtension({
  name: PRESERVED_BLOCK_NODE_NAME,
  schemaNodeName: PRESERVED_BLOCK_NODE_NAME,
  nodeSpec: {
    group: "block",
    atom: true,
    selectable: false,
    attrs: {
      xml: {},
    },
    parseDOM: [
      {
        tag: "div[data-docx-preserved-block]",
        getAttrs(dom) {
          if (!(dom instanceof HTMLElement)) {
            return false;
          }
          const xml = dom.dataset["docxPreservedBlock"];
          return xml === undefined || xml === "" ? false : { xml };
        },
      },
    ],
    toDOM(node) {
      const { xml } = expectPreservedBlockAttrs(node);
      return [
        "div",
        {
          "data-docx-preserved-block": xml,
          "aria-hidden": "true",
          contenteditable: "false",
          style: "display: none;",
        },
      ];
    },
  },
  onSchemaReady: () => ({
    plugins: [
      new Plugin({
        appendTransaction(transactions, _oldState, newState) {
          if (!transactions.some(({ docChanged }) => docChanged)) {
            return null;
          }
          const unwritable = collectUnwritableBlocks(newState.doc);
          if (unwritable.length === 0) {
            return null;
          }
          const transaction = newState.tr;
          for (const { node, position } of unwritable.toSorted(
            (first, second) => second.position - first.position,
          )) {
            transaction.delete(position, position + node.nodeSize);
          }
          return transaction;
        },
      }),
    ],
  }),
});
