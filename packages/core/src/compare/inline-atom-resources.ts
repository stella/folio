import type { Node as PMNode } from "prosemirror-model";

import { rebindDrawingImageRelationship } from "../docx/drawingRelationships";
import { expectImageAttrs } from "../prosemirror/attrs";

/** Detach supported atoms from the source package; save allocates the receiving media relationship. */
export const prepareTargetInlineAtom = (node: PMNode): PMNode | null => {
  if (node.type.name === "field") return node;
  if (node.type.name === "pageBreakRun") return node;
  if (node.type.name !== "image") return null;
  const attrs = expectImageAttrs(node);
  if (!attrs.src.startsWith("data:") || attrs.hlinkRId || attrs.hlinkHref) return null;
  const rId = "rId_img_compare";
  const rawXml = attrs._docxRawXml;
  const reboundXml = rawXml
    ? rebindDrawingImageRelationship({ xml: rawXml, previousId: attrs.rId ?? "", nextId: rId })
    : undefined;
  if (reboundXml === null) return null;
  return node.type.create(
    { ...node.attrs, rId, ...(reboundXml !== undefined && { _docxRawXml: reboundXml }) },
    node.content,
    node.marks,
  );
};
