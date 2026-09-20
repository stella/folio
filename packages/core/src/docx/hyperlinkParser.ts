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
  withPreservedChildren,
} from "./containerChildren";
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
  // Whether to add to browser history
  if (parseOnOffAttribute(node, "w", "history") === true) {
    hyperlink.history = true;
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
 * Two callers read this one map: {@link parseHyperlink}, and the paragraph
 * parser's revision-segmenting walk, which overrides the four
 * `CT_RunTrackChange` wrappers because OOXML nests a revision inside a link
 * and the model nests the link inside the revision. Everything else is
 * decided here once, so a child one caller starts recognising is recognised
 * by both rather than by whichever list somebody remembered to update.
 */
export const hyperlinkChildHandlers = ({
  push,
  styles,
  theme,
  rels,
  media,
  inScopeXmlns,
}: HyperlinkChildContext): ChildHandlers<"w:hyperlink"> => ({
  // A link's run answers the same keep question a paragraph's run does. No
  // pass fills a run inside a link later, so the question is the model's
  // alone, and a run that holds nothing is not admitted: the save would write
  // a run the next parse drops.
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

  // Transparent wrappers. The markup stays opaque, but the text it wraps is
  // on the line, so the capture carries it and a linked party name still
  // reads in the editor.
  customXml: (child) => {
    push(preserveInlineChild(child));
  },
  smartTag: (child) => {
    push(preserveInlineChild(child));
  },

  // A revision *inside* a link. The paragraph parser hoists these around the
  // link instead; a link reached from anywhere else — a simple field's cached
  // result — keeps the markup rather than dropping it.
  del: CAPTURE,
  ins: CAPTURE,
  moveFrom: CAPTURE,
  moveTo: CAPTURE,

  bdo: CAPTURE,
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
  dir: CAPTURE,
  fldSimple: CAPTURE,
  hyperlink: CAPTURE,
  moveFromRangeEnd: CAPTURE,
  moveFromRangeStart: CAPTURE,
  moveToRangeEnd: CAPTURE,
  moveToRangeStart: CAPTURE,
  permEnd: CAPTURE,
  permStart: CAPTURE,
  proofErr: CAPTURE,
  sdt: CAPTURE,
  subDoc: CAPTURE,
});

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

  for (const child of hyperlink.children) {
    if (child.type === "run") {
      for (const content of child.content) {
        if (content.type === "text") {
          text += content.text;
        } else if (content.type === "tab") {
          text += "\t";
        }
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
  return hyperlink.children.some((child) => child.type === "run");
}

/**
 * Get all runs from a hyperlink
 *
 * @param hyperlink - Parsed Hyperlink object
 * @returns Array of Run objects
 */
export function getHyperlinkRuns(hyperlink: Hyperlink): Run[] {
  return hyperlink.children.filter((child): child is Run => child.type === "run");
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
