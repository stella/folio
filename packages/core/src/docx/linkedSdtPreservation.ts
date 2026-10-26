import { TaggedError, type TaggedErrorClass } from "better-result";

import type { PreservedInline } from "../types/document";
import { preserveInlineChild } from "./preservedRunContent";
import {
  getChildElements,
  getLocalName,
  WORDPROCESSINGML_NAMESPACE_URIS,
  type XmlElement,
} from "./xmlParser";

const UnrepresentableLinkedSdtRevisionErrorBase: TaggedErrorClass<"UnrepresentableLinkedSdtRevisionError"> =
  TaggedError("UnrepresentableLinkedSdtRevisionError");

export class UnrepresentableLinkedSdtRevisionError extends UnrepresentableLinkedSdtRevisionErrorBase<{
  message: string;
}> {}

/** Capture only the unsupported hyperlink child, never its enclosing revision. */
export const preserveLinkedSdt = (root: XmlElement): PreservedInline => {
  const pending = getChildElements(root);
  while (pending.length > 0) {
    const node = pending.pop();
    if (!node) continue;
    const wordprocessing = WORDPROCESSINGML_NAMESPACE_URIS.has(
      node.namespaceUri ?? "",
    );
    const name = getLocalName(node.name);
    if (
      wordprocessing &&
      [
        "ins",
        "del",
        "moveFrom",
        "moveTo",
        "rPrChange",
        "pPrChange",
        "sectPrChange",
        "tblPrChange",
        "tblGridChange",
        "trPrChange",
        "tblPrExChange",
        "tcPrChange",
        "cellIns",
        "cellDel",
        "cellMerge",
        "numberingChange",
      ].includes(name)
    ) {
      throw new UnrepresentableLinkedSdtRevisionError({
        message:
          "A content control under a hyperlink cannot preserve nested revisions as opaque XML.",
      });
    }
    pending.push(...getChildElements(node));
  }
  return preserveInlineChild(root);
};
