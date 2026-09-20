import type { Node as PMNode } from "prosemirror-model";

import { type BuiltInStyleIndex, resolveHeadingLevel } from "../docx/builtInStyles";
import { readOutlineLevelAttr } from "../prosemirror/outlineLevelAttr";

/**
 * Information about a heading found in the document.
 */
export type HeadingInfo = {
  /** The text content of the heading */
  text: string;
  /** Outline level (0 = Heading 1, 1 = Heading 2, etc.) */
  level: number;
  /** ProseMirror document position of the paragraph node */
  pmPos: number;
  /** 1-indexed page number, filled in by the editor once layout is known. */
  pageNumber?: number | null;
};

/**
 * Collect all headings from a ProseMirror document, in document order.
 *
 * Classification is {@link resolveHeadingLevel}'s: the paragraph's effective
 * outline level, else the style's built-in `w:name`. Pass the open document's
 * index (`styleResolver.builtInStyles`) so a localized style id such as
 * `Nadpis1` or `berschrift1` resolves through its name rather than falling out
 * of the outline.
 */
export function collectHeadings(doc: PMNode, styles: BuiltInStyleIndex): HeadingInfo[] {
  const headings: HeadingInfo[] = [];

  doc.descendants((node, pos) => {
    if (node.type.name !== "paragraph") {
      return;
    }
    const level = resolveHeadingLevel(
      {
        outlineLevel: readOutlineLevelAttr(node.attrs["outlineLevel"]),
        styleId: typeof node.attrs["styleId"] === "string" ? node.attrs["styleId"] : null,
      },
      styles,
    );
    if (level === undefined) {
      return;
    }
    let text = "";
    // oxlint-disable-next-line unicorn/no-array-for-each -- ProseMirror Node.forEach
    node.forEach((child) => {
      if (child.isText) {
        text += child.text || "";
      }
    });
    if (text.trim()) {
      headings.push({ text: text.trim(), level, pmPos: pos });
    }
  });

  return headings;
}
