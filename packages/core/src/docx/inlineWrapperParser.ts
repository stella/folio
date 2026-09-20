/**
 * The four transparent inline wrappers, read from the element that spells one.
 *
 * `w:bdo`, `w:dir`, `w:smartTag` and the run-level `w:customXml` are declared
 * by `EG_PContent`, so every container that holds inline content holds them:
 * a paragraph, a revision, an inline content control, a link and a simple
 * field. What the wrapper adds — a layout control, or a name, a URI and a
 * properties bag — is decided by the element alone, so it is read here once.
 *
 * What the wrapper *holds* is the container's own question, because the
 * declared children of a `w:bdo` are the declared children of whatever the
 * `w:bdo` sits in. So the caller walks the content with its own handler map
 * and hands the result in; this module never recurses.
 */

import type { InlineWrapper, ParagraphContent } from "../types/document";

import { BIDI_CONTROLS } from "@stll/docx-core/model";
import type { BidiControl } from "@stll/docx-core/model";

import { captureVerbatimXml } from "./verbatimCapture";
import { findChild, getAttribute } from "./xmlParser";
import type { XmlElement } from "./xmlParser";

/** The elements that spell a transparent inline wrapper. */
export const INLINE_WRAPPER_ELEMENTS = ["bdo", "dir", "smartTag", "customXml"] as const;

/** One of them, as the child dispatcher names it. */
export type InlineWrapperElement = (typeof INLINE_WRAPPER_ELEMENTS)[number];

/** The wrapper kinds that name an element rather than a layout control. */
type TaggedInlineWrapper = Extract<InlineWrapper, { element: string }>;

/**
 * A `w:bdo` or `w:dir`: a bidirectional override or embedding. Both hold
 * inline content and change only how it is laid out, so the wrapper carries
 * its direction and nothing else.
 */
const bidiWrapper = (
  node: XmlElement,
  control: BidiControl,
  content: ParagraphContent[],
): InlineWrapper => {
  const wrapper: InlineWrapper = {
    type: "inlineWrapper",
    kind: "bidi",
    control,
    content,
  };
  const direction = getAttribute(node, "w", "val");
  if (direction === "ltr" || direction === "rtl") {
    wrapper.direction = direction;
  }
  return wrapper;
};

/**
 * A `w:smartTag` or a run-level `w:customXml`. Both name an element in another
 * vocabulary around content that is ordinary inline content; what the wrapper
 * adds is its name, its namespace and a properties bag folio replays rather
 * than reads. The `w:element` attribute is required by both content models,
 * and a tag without one names nothing, so it is read as the empty string
 * rather than refused: the content is what matters.
 */
const taggedWrapper = (
  node: XmlElement,
  kind: TaggedInlineWrapper["kind"],
  content: ParagraphContent[],
): TaggedInlineWrapper => {
  const wrapper: TaggedInlineWrapper = {
    type: "inlineWrapper",
    kind,
    element: getAttribute(node, "w", "element") ?? "",
    content,
  };
  const uri = getAttribute(node, "w", "uri");
  if (uri !== null) {
    wrapper.uri = uri;
  }
  const properties = findChild(node, "w", kind === "smartTag" ? "smartTagPr" : "customXmlPr");
  if (properties) {
    wrapper.propertiesXml = captureVerbatimXml(properties);
  }
  return wrapper;
};

/** The wrapper `element` spells, around content the caller has already read. */
export const inlineWrapperOf = (
  element: InlineWrapperElement,
  node: XmlElement,
  content: ParagraphContent[],
): InlineWrapper => {
  switch (element) {
    case "bdo":
      return bidiWrapper(node, BIDI_CONTROLS.override, content);
    case "dir":
      return bidiWrapper(node, BIDI_CONTROLS.embedding, content);
    case "smartTag":
      return taggedWrapper(node, "smartTag", content);
    case "customXml":
      return taggedWrapper(node, "customXml", content);
    default: {
      const unread: never = element;
      return unread;
    }
  }
};
