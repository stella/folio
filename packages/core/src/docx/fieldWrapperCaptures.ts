/**
 * A simple field's transparent wrappers, holding only what the field holds.
 *
 * The model types a wrapper's content as any paragraph content, so a field
 * built through the model may hold, inside a `w:dir` or a smart tag, an
 * equation, a bookmark, a content control or another field. The parser never
 * reads one: a wrapper inside `w:fldSimple` reads with the field's own map,
 * which keeps everything else as the markup it is. The editor's field node
 * holds the field's content only, so a built field reads here as the parser
 * would read the markup the serializer writes for it.
 */

import { panic } from "better-result";

import type {
  InlineWrapper,
  ParagraphContent,
  SimpleField,
} from "../types/document";
import {
  isHyperlinkContent,
  isInlineSdtContent,
  isSimpleFieldContent,
  isTrackedChangeWrapperChild,
} from "./inlineWrapperContent";
import { preservedInlineCapture } from "./preservedRunContent";
import { serializeParagraph } from "./serializer/paragraphSerializer";

/** The markup the serializer writes for one inline item. */
const writtenMarkup = (item: ParagraphContent): string => {
  if (item.type === "mathEquation") return item.ommlXml;
  const written = serializeParagraph({ type: "paragraph", content: [item] });
  const open = "<w:p>";
  const close = "</w:p>";
  if (!written.startsWith(open) || !written.endsWith(close))
    return panic(
      `A lone ${item.type} must serialize as bare paragraph content`,
    );
  return written.slice(open.length, -close.length);
};

/** Comment range markers a field's conversion reads where they stand. */
const readInField = (item: ParagraphContent): boolean =>
  isSimpleFieldContent(item) ||
  item.type === "commentRangeStart" ||
  item.type === "commentRangeEnd";

const wrapperInField = (wrapper: InlineWrapper): InlineWrapper => {
  let changed = false;
  const content = wrapper.content.map((item): ParagraphContent => {
    if (item.type === "inlineWrapper") {
      const inner = wrapperInField(item);
      if (inner !== item) changed = true;
      return inner;
    }
    if (readInField(item)) {
      const next = normalizedItem(item);
      if (next !== item) changed = true;
      return next;
    }
    changed = true;
    return preservedInlineCapture(writtenMarkup(item));
  });
  return changed ? { ...wrapper, content } : wrapper;
};

const fieldWithCaptures = (field: SimpleField): SimpleField => {
  let changed = false;
  const content = field.content.map((item) => {
    const next =
      item.type === "inlineWrapper"
        ? wrapperInField(item)
        : withFieldWrapperCaptures([item]).at(0);
    if (next === undefined || !isSimpleFieldContent(next))
      return panic("A field's content must stay field content");
    if (next !== item) changed = true;
    return next;
  });
  return changed ? { ...field, content } : field;
};

/**
 * Inline content with every simple field's wrappers holding only field
 * content: anything else they hold becomes the capture the parser would read.
 * Returns `content` itself when no field needed it.
 */
export const withFieldWrapperCaptures = (
  content: readonly ParagraphContent[],
): readonly ParagraphContent[] => {
  let changed = false;
  const next = content.map((item): ParagraphContent => {
    const normalized = normalizedItem(item);
    if (normalized !== item) changed = true;
    return normalized;
  });
  return changed ? next : content;
};

const normalizedItem = (item: ParagraphContent): ParagraphContent => {
  switch (item.type) {
    case "simpleField":
      return fieldWithCaptures(item);
    case "hyperlink": {
      const children = withFieldWrapperCaptures(item.children);
      if (children === item.children) return item;
      // Each child stays what it was; only a field inside one changes.
      const linked = children.filter(isHyperlinkContent);
      if (linked.length !== children.length)
        return panic("A link's children must stay link content");
      return { ...item, children: linked };
    }
    case "inlineWrapper": {
      const inner = withFieldWrapperCaptures(item.content);
      return inner === item.content ? item : { ...item, content: [...inner] };
    }
    case "insertion":
    case "deletion":
    case "moveFrom":
    case "moveTo": {
      const inner = withFieldWrapperCaptures(item.content);
      return inner === item.content
        ? item
        : { ...item, content: inner.filter(isTrackedChangeWrapperChild) };
    }
    case "inlineSdt": {
      const inner = withFieldWrapperCaptures(item.content);
      return inner === item.content
        ? item
        : { ...item, content: inner.filter(isInlineSdtContent) };
    }
    default:
      return item;
  }
};
