/**
 * Hyperlink Parser - Parse hyperlinks (w:hyperlink) with URL resolution
 *
 * OOXML Reference:
 * - Hyperlink element: w:hyperlink
 * - Attributes:
 *   - r:id - Relationship ID for external link (resolves via .rels)
 *   - w:anchor - Internal bookmark name
 *   - w:tooltip - Tooltip/title text
 *   - w:tgtFrame - Target frame (_blank, _self, etc.)
 *   - w:history - Whether to add to history
 *   - w:docLocation - Location within a document
 *
 * External links use r:id to reference a relationship in document.xml.rels
 * Internal links use w:anchor to reference a bookmark in the same document
 */

import type {
  Hyperlink,
  InlineWrapper,
  ParagraphContent,
  Run,
  BookmarkStart,
  BookmarkEnd,
  Theme,
  RelationshipMap,
  MediaFile,
} from "../types/document";

import { sanitizeExternalUrl } from "../utils/urlSecurity";
import {
  CAPTURE,
  type ChildHandlers,
  dispatchChildren,
  ownedElsewhere,
  withPreservedChildren,
} from "./containerChildren";
import { inlineWrapperOf } from "./inlineWrapperParser";
import type { InlineWrapperElement } from "./inlineWrapperParser";
import { preservedInlineCapture, preserveInlineChild } from "./preservedRunContent";
import { RELATIONSHIP_TYPES, resolveRelationshipIdOfType } from "./relsParser";
import { parseRun } from "./runParser";
import { runHoldsPayload } from "./runPayload";
import type { StyleMap } from "./styleParser";
import {
  getAttribute,
  mergeXmlnsDeclarations,
  parseNumericAttribute,
  parseOnOffAttribute,
} from "./xmlParser";
import type { XmlElement } from "./xmlParser";

// ============================================================================
// HYPERLINK PARSER
// ============================================================================

/**
 * Parse bookmark start (w:bookmarkStart)
 *
 * Used for internal hyperlink targets within the document.
 */
function parseBookmarkStart(node: XmlElement): BookmarkStart {
  const id = parseNumericAttribute(node, "w", "id") ?? 0;
  const name = getAttribute(node, "w", "name") ?? "";

  const bookmark: BookmarkStart = {
    type: "bookmarkStart",
    id,
    name,
  };

  // Table column bookmarks
  const colFirst = parseNumericAttribute(node, "w", "colFirst");
  if (colFirst !== undefined) {
    bookmark.colFirst = colFirst;
  }

  const colLast = parseNumericAttribute(node, "w", "colLast");
  if (colLast !== undefined) {
    bookmark.colLast = colLast;
  }

  return bookmark;
}

/**
 * Parse bookmark end (w:bookmarkEnd)
 */
function parseBookmarkEnd(node: XmlElement): BookmarkEnd {
  const id = parseNumericAttribute(node, "w", "id") ?? 0;

  return {
    type: "bookmarkEnd",
    id,
  };
}

/**
 * Parse a hyperlink element (w:hyperlink)
 *
 * Handles both external links (via r:id relationship) and internal
 * links (via w:anchor bookmark reference).
 *
 * @param node - The w:hyperlink XML element
 * @param rels - Relationship map to resolve r:id references
 * @param styles - Style map for resolving run styles
 * @param theme - Theme for resolving colors/fonts
 * @param media - Media files map for image data
 * @returns Parsed Hyperlink object
 */
export function parseHyperlink(
  node: XmlElement,
  rels: RelationshipMap | null,
  styles: StyleMap | null = null,
  theme: Theme | null = null,
  media: Map<string, MediaFile> | null = null,
  rootXmlns: Record<string, string> = {},
): Hyperlink {
  const hyperlink: Hyperlink = {
    type: "hyperlink",
    children: [],
  };

  // === External Link (r:id) ===
  // Get relationship ID for external links
  const rId = getAttribute(node, "r", "id");
  if (rId) {
    hyperlink.rId = rId;

    // Resolve the relationship to get the actual URL. An id that names some
    // other kind of part names no URL: reading its target anyway would turn a
    // broken link into a link to a package part. An id that resolves to
    // nothing is still written back, and the parse reports those counts once
    // it has the whole body (`countDanglingRelationshipReferences`).
    const resolved = resolveRelationshipIdOfType(rels, rId, RELATIONSHIP_TYPES.hyperlink);
    if (resolved.status === "resolved") {
      // External hyperlinks have TargetMode="External" and target is the URL
      // Both external and internal links use the same target
      const safeHref = sanitizeExternalUrl(resolved.relationship.target);
      if (safeHref) {
        // Validate the protocol without rewriting the authored relationship target.
        hyperlink.href = resolved.relationship.target;
      }
    }
  }

  // === Internal Bookmark Link (w:anchor) ===
  // Get internal bookmark anchor
  const anchor = getAttribute(node, "w", "anchor");
  if (anchor) {
    hyperlink.anchor = anchor;
    // For internal links, create a fragment-style href
    if (!hyperlink.href) {
      hyperlink.href = `#${anchor}`;
    }
  }

  // === Tooltip ===
  const tooltip = getAttribute(node, "w", "tooltip");
  if (tooltip) {
    hyperlink.tooltip = tooltip;
  }

  // === Target Frame ===
  // The authored frame name, verbatim: a document says what it says, and a save
  // must be able to write it back. Common values are _blank, _self, _parent and
  // _top, but a named frame is legal and was silently rewritten to _blank here.
  // The allow-list clamp belongs where a DOM target is produced
  // (`anchorTargetAttrs`), not where the package is read.
  const tgtFrame = getAttribute(node, "w", "tgtFrame");
  if (tgtFrame) {
    hyperlink.target = tgtFrame;
  }

  // === History ===
  // Whether to add to browser history. Three states: an explicit `w:history="0"`
  // is not the attribute being absent, and the serializer already writes each
  // back, so reading only the on left its off branch unreachable.
  const history = parseOnOffAttribute(node, "w", "history");
  if (history !== undefined) {
    hyperlink.history = history;
  }

  // === Document Location ===
  // Location within a linked document (like fragment for external doc)
  const docLocation = getAttribute(node, "w", "docLocation");
  if (docLocation) {
    hyperlink.docLocation = docLocation;
  }

  // === Parse Children ===
  // Accumulate the hyperlink's own xmlns onto the inherited set so a captured
  // VML `w:pict` inside a run resolves a non-canonical prefix scoped on the
  // `w:hyperlink` wrapper itself.
  const inScopeXmlns = mergeXmlnsDeclarations(rootXmlns, node);
  const children: Hyperlink["children"] = [];
  const preserved = dispatchChildren({
    element: node,
    container: "w:hyperlink",
    capturePosition: () => children.length,
    handlers: hyperlinkChildHandlers({
      push: (child) => {
        children.push(child);
      },
      styles,
      theme,
      rels,
      media,
      inScopeXmlns,
    }),
  });
  hyperlink.children = withPreservedChildren(children, preserved, preservedInlineCapture);

  return hyperlink;
}

/** What {@link hyperlinkChildHandlers} needs to read one child of a link. */
export type HyperlinkChildContext = {
  /** Where a parsed or captured child lands, in source order. */
  push: (child: Hyperlink["children"][number]) => void;
  styles: StyleMap | null;
  theme: Theme | null;
  rels: RelationshipMap | null;
  media: Map<string, MediaFile> | null;
  inScopeXmlns: Record<string, string>;
};

/**
 * What a `w:hyperlink` does with every child its content model declares.
 *
 * `CT_Hyperlink` is `EG_PContent`, so a link may hold a permission range, a
 * proofing error, a nested field or one of the eight custom-XML revision
 * ranges between its runs; folio models three of the thirty-two names and
 * used to drop the other twenty-nine off the end of a `switch`.
 *
 * Three callers read this one map: {@link parseHyperlink}, the paragraph
 * parser's revision-segmenting walk, which overrides the four
 * `CT_RunTrackChange` wrappers because OOXML nests a revision inside a link
 * and the model nests the link inside the revision, and
 * {@link parseLinkedInlineWrapper}, which walks a transparent wrapper the link
 * holds. Everything else is decided here once, so a child one caller starts
 * recognising is recognised by all of them rather than by whichever list
 * somebody remembered to update.
 */
export const hyperlinkChildHandlers = (context: HyperlinkChildContext) => {
  const { push, styles, theme, rels, media, inScopeXmlns } = context;
  const wrapper =
    (element: InlineWrapperElement) =>
    (child: XmlElement): void => {
      push(parseLinkedInlineWrapper(element, child, context));
    };
  return {
    r: (child) => {
      const run = parseRun(child, styles, theme, rels, media, inScopeXmlns);
      if (runHoldsPayload(run)) {
        push(run);
      }
    },
    bookmarkStart: (child) => {
      push(parseBookmarkStart(child));
    },
    bookmarkEnd: (child) => {
      push(parseBookmarkEnd(child));
    },

    // The transparent wrappers. `EG_PContent` declares all four, so a link may
    // be authored around a bidirectional override or a smart tag; each is read
    // as the wrapper it is, and the runs inside it stay editable text.
    bdo: wrapper("bdo"),
    dir: wrapper("dir"),
    customXml: wrapper("customXml"),
    smartTag: wrapper("smartTag"),

    // A revision *inside* a link. The paragraph parser hoists these around the
    // link instead; a link reached from anywhere else — a simple field's cached
    // result — keeps the markup rather than dropping it.
    del: CAPTURE,
    ins: CAPTURE,
    moveFrom: CAPTURE,
    moveTo: CAPTURE,

    commentRangeEnd: CAPTURE,
    commentRangeStart: CAPTURE,
    customXmlDelRangeEnd: CAPTURE,
    customXmlDelRangeStart: CAPTURE,
    customXmlInsRangeEnd: CAPTURE,
    customXmlInsRangeStart: CAPTURE,
    customXmlMoveFromRangeEnd: CAPTURE,
    customXmlMoveFromRangeStart: CAPTURE,
    customXmlMoveToRangeEnd: CAPTURE,
    customXmlMoveToRangeStart: CAPTURE,
    fldSimple: (child) => {
      push(preserveInlineChild(child));
    },
    hyperlink: (child) => {
      push(preserveInlineChild(child));
    },
    moveFromRangeEnd: CAPTURE,
    moveFromRangeStart: CAPTURE,
    moveToRangeEnd: CAPTURE,
    moveToRangeStart: CAPTURE,
    permEnd: CAPTURE,
    permStart: CAPTURE,
    proofErr: CAPTURE,
    sdt: CAPTURE,
    subDoc: CAPTURE,
  } satisfies ChildHandlers<"w:hyperlink">;
};

const LINKED_SMART_TAG_PROPERTIES_OWNER = ownedElsewhere({
  container: "run-level-content",
  child: "smartTagPr",
  reader: "inlineWrapperParser#inlineWrapperOf",
});

const LINKED_CUSTOM_XML_PROPERTIES_OWNER = ownedElsewhere({
  container: "run-level-content",
  child: "customXmlPr",
  reader: "inlineWrapperParser#inlineWrapperOf",
});

/**
 * A transparent wrapper a link holds, with the content the link would hold.
 *
 * `CT_BdoContentRun` and its three siblings are `EG_PContent`, the group
 * `CT_Hyperlink` is, so the wrapper's declared children are run-level content
 * and the decision per child is the link's own map. That is what keeps the two
 * from drifting: a `w:ins` inside a `w:bdo` inside a link is captured for the
 * same reason a `w:ins` directly inside the link is.
 *
 * The wrapper's own `w:smartTagPr` / `w:customXmlPr` is read by
 * {@link inlineWrapperOf}, so it is declared owned here rather than captured a
 * second time by the sink.
 */
const parseLinkedInlineWrapper = (
  element: InlineWrapperElement,
  node: XmlElement,
  context: HyperlinkChildContext,
): InlineWrapper => {
  const inScopeXmlns = mergeXmlnsDeclarations(context.inScopeXmlns, node);
  const content: Hyperlink["children"] = [];
  const preserved = dispatchChildren({
    element: node,
    container: "run-level-content",
    capturePosition: () => content.length,
    handlers: {
      ...hyperlinkChildHandlers({
        ...context,
        inScopeXmlns,
        push: (child) => {
          content.push(child);
        },
      }),
      customXmlPr: LINKED_CUSTOM_XML_PROPERTIES_OWNER,
      smartTagPr: LINKED_SMART_TAG_PROPERTIES_OWNER,
      // A `w:pPr` is not a child of any of the four wrappers; the declared set
      // is shared with `w:p`, where the paragraph reads it off the element.
      pPr: CAPTURE,
    },
  });
  return inlineWrapperOf(
    element,
    node,
    withPreservedChildren(content, preserved, preservedInlineCapture),
  );
};

// ============================================================================
// UTILITY FUNCTIONS
// ============================================================================

/**
 * Get the display text of a hyperlink
 *
 * Concatenates text from all child runs.
 *
 * @param hyperlink - Parsed Hyperlink object
 * @returns Display text string
 */
export function getHyperlinkText(hyperlink: Hyperlink): string {
  let text = "";

  for (const run of getHyperlinkRuns(hyperlink)) {
    for (const content of run.content) {
      if (content.type === "text") {
        text += content.text;
      } else if (content.type === "tab") {
        text += "\t";
      }
    }
  }

  return text;
}

/**
 * Check if a hyperlink is an external link
 *
 * @param hyperlink - Parsed Hyperlink object
 * @returns true if this links to an external URL
 */
export function isExternalLink(hyperlink: Hyperlink): boolean {
  // Has rId and resolved href that starts with a protocol
  if (hyperlink.href) {
    return sanitizeExternalUrl(hyperlink.href) !== undefined;
  }
  // Has rId but not resolved (still counts as external attempt)
  return !!hyperlink.rId && !hyperlink.anchor;
}

/**
 * Check if a hyperlink is an internal bookmark link
 *
 * @param hyperlink - Parsed Hyperlink object
 * @returns true if this links to an internal bookmark
 */
export function isInternalLink(hyperlink: Hyperlink): boolean {
  return !!hyperlink.anchor;
}

/**
 * Get the resolved URL of a hyperlink
 *
 * For external links, returns the full URL.
 * For internal links, returns the anchor prefixed with #.
 * Returns undefined if the link couldn't be resolved.
 *
 * @param hyperlink - Parsed Hyperlink object
 * @returns Resolved URL or undefined
 */
export function getHyperlinkUrl(hyperlink: Hyperlink): string | undefined {
  return hyperlink.href;
}

/**
 * Check if a hyperlink has any content (runs)
 *
 * @param hyperlink - Parsed Hyperlink object
 * @returns true if hyperlink has child runs
 */
export function hasContent(hyperlink: Hyperlink): boolean {
  return getHyperlinkRuns(hyperlink).length > 0;
}

/**
 * Get all runs from a hyperlink, including those inside a transparent wrapper
 *
 * A `w:bdo`, `w:dir`, `w:smartTag` or run-level `w:customXml` says how the
 * linked text is laid out or what it is tagged as, never that it is not the
 * link's text, so the walk goes through it.
 *
 * @param hyperlink - Parsed Hyperlink object
 * @returns Array of Run objects
 */
export function getHyperlinkRuns(hyperlink: Hyperlink): Run[] {
  const runs: Run[] = [];
  const collect = (items: readonly ParagraphContent[]): void => {
    for (const item of items) {
      if (item.type === "run") {
        runs.push(item);
        continue;
      }
      if (item.type === "inlineWrapper") {
        collect(item.content);
      }
    }
  };
  collect(hyperlink.children);
  return runs;
}

/**
 * Resolve a hyperlink's rId to a URL using a relationship map
 *
 * This is useful when you have a hyperlink that was parsed without
 * relationship context and need to resolve it later.
 *
 * @param hyperlink - Parsed Hyperlink object (will be modified)
 * @param rels - Relationship map to resolve against
 * @returns The resolved URL or undefined
 */
export function resolveHyperlinkUrl(
  hyperlink: Hyperlink,
  rels: RelationshipMap,
): string | undefined {
  const resolved = resolveRelationshipIdOfType(rels, hyperlink.rId, RELATIONSHIP_TYPES.hyperlink);
  if (resolved.status === "resolved") {
    const safeHref = sanitizeExternalUrl(resolved.relationship.target);
    if (safeHref) {
      hyperlink.href = resolved.relationship.target;
      return hyperlink.href;
    }
  }

  // If there's an anchor but no href yet, set it
  if (hyperlink.anchor && !hyperlink.href) {
    hyperlink.href = `#${hyperlink.anchor}`;
    return hyperlink.href;
  }

  return hyperlink.href;
}

/**
 * Create an internal hyperlink to a bookmark
 *
 * Utility function for creating hyperlinks programmatically.
 *
 * @param anchor - Bookmark name to link to
 * @param children - Child runs for display text
 * @param options - Optional properties (tooltip, etc.)
 * @returns New Hyperlink object
 */
export function createInternalHyperlink(
  anchor: string,
  children: Run[],
  options?: {
    tooltip?: string;
  },
): Hyperlink {
  return {
    type: "hyperlink",
    anchor,
    href: `#${anchor}`,
    ...(options?.tooltip !== undefined ? { tooltip: options.tooltip } : {}),
    children,
  };
}

/**
 * Create an external hyperlink
 *
 * Utility function for creating hyperlinks programmatically.
 * Note: The rId would need to be assigned when serializing.
 *
 * @param url - External URL
 * @param children - Child runs for display text
 * @param options - Optional properties (tooltip, target, etc.)
 * @returns New Hyperlink object
 */
export function createExternalHyperlink(
  url: string,
  children: Run[],
  options?: {
    tooltip?: string;
    target?: string;
  },
): Hyperlink {
  return {
    type: "hyperlink",
    href: url,
    ...(options?.tooltip !== undefined ? { tooltip: options.tooltip } : {}),
    ...(options?.target !== undefined ? { target: options.target } : {}),
    children,
  };
}
