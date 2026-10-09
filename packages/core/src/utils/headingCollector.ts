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

const RUN_IN_TITLE_MARKS = new Set(["bold", "italic", "underline"]);

type RunInTitleCapture =
  | { status: "searching" }
  | { status: "collecting"; format: string; text: string }
  | { status: "complete"; text: string };

const runInTitleFormat = (child: PMNode): string | undefined => {
  const marks = child.marks
    .map(({ type }) => type.name)
    .filter((name) => RUN_IN_TITLE_MARKS.has(name))
    .sort();
  return marks.length > 0 ? marks.join("+") : undefined;
};

const headingText = (paragraph: PMNode): string => {
  let text = "";
  let capture: RunInTitleCapture = { status: "searching" };

  // ProseMirror can split one authored run into adjacent text nodes when
  // comment or run-identity marks differ, so collect the emphasized span by
  // formatting rather than treating every text node as a DOCX run.
  for (let childIndex = 0; childIndex < paragraph.childCount; childIndex += 1) {
    const child = paragraph.child(childIndex);
    if (!child.isText) {
      if (capture.status === "collecting") {
        capture = { status: "complete", text: capture.text };
      }
      continue;
    }

    const childText = child.text || "";
    text += childText;
    const format = runInTitleFormat(child);
    if (capture.status === "searching" && format !== undefined) {
      capture = { status: "collecting", format, text: childText };
      continue;
    }
    if (capture.status !== "collecting") {
      continue;
    }
    capture =
      format === capture.format
        ? { status: "collecting", format: capture.format, text: capture.text + childText }
        : { status: "complete", text: capture.text };
  }

  const formattedText = capture.status === "searching" ? "" : capture.text.trim();
  return (formattedText || text).trim();
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
    const text = headingText(node);
    if (text.trim()) {
      headings.push({ text: text.trim(), level, pmPos: pos });
    }
  });

  return headings;
}
