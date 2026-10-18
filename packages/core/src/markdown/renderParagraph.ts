/**
 * Render a single paragraph as a block of markdown. Three cases: heading style
 * → `#`…`######`; list item → indented marker + inline content (Word's exact
 * marker preserved); plain prose → escaped inline content. Word's `Quote` /
 * `IntenseQuote` styles become blockquotes. Ported from eigenpal/docx-editor
 * PR #595.
 *
 * A numbered heading keeps its number after the hashes (`## 1. Scope`).
 *
 * Which paragraphs those are is `builtInStyles`' decision, not this file's: the
 * id `Nadpis1` and the id `Heading1` are the same heading, and `ClauseHeading1`
 * is not one at all.
 */

import { isQuoteStyle, resolveHeadingLevel } from "../docx/builtInStyles";
import { getCachedNumberingMap } from "../docx/numberingParser";
import { listLabelAttrsFromRendering } from "../prosemirror/listLabels";
import { resolveParagraphList } from "../prosemirror/paragraphList";
import type { DocxPackage, ListRendering, Paragraph } from "../types/document";
import { renderParagraphInline } from "./renderRuns";
import { inlineBreaksToHtml } from "./escape";
import type { RenderContext } from "./types";

/** Markdown has six heading levels; Word has nine. */
const MAX_MARKDOWN_HEADING_LEVEL = 6;

/**
 * Render a paragraph and return the block text. No surrounding blank line: the
 * caller joins blocks.
 */
export function renderParagraph(
  ctx: RenderContext,
  pkg: DocxPackage | undefined,
  para: Paragraph,
): string {
  return renderParagraphBlock(ctx, pkg, para).markdown;
}

/** List items declare the boundary needed to preserve their Markdown structure. */
export type RenderedParagraph =
  | { type: "paragraph"; markdown: string }
  | { type: "listItem"; markdown: string; separation: "tight" | "loose" };

/**
 * Render a paragraph. A list item is a paragraph that shows a marker, which is
 * when the list counter gives it a label: a numbered heading renders as a
 * heading, and a level that hides its marker (`w:vanish`) or that its list
 * does not define as prose, as every other reader reads them.
 */
export function renderParagraphBlock(
  ctx: RenderContext,
  pkg: DocxPackage | undefined,
  para: Paragraph,
): RenderedParagraph {
  // Every paragraph advances the list counter, whatever it renders as: the
  // items after a numbered heading continue from its number.
  const inherited = ctx.paragraphNumbering(para.formatting?.styleId);
  const numPrFromStyle = inherited ?? para.formatting?.numPrFromStyle;
  const numbering = pkg?.numbering ? getCachedNumberingMap(pkg.numbering) : null;
  const resolvedList = resolveParagraphList(
    { numPr: para.formatting?.numPr, numPrFromStyle },
    numbering,
  );
  const resolvedRendering = resolvedList.kind === "marker" ? resolvedList.rendering : undefined;
  const counterAttrs = listLabelAttrsFromRendering(resolvedRendering, {
    numPr: para.formatting?.numPr,
    numPrFromStyle,
  });
  if (para.pPrMark) counterAttrs.pPrMark = para.pPrMark;
  let originalList: ListRendering | undefined;
  if (para.propertyChanges) {
    counterAttrs._propertyChanges = para.propertyChanges.map(
      ({ previousFormatting, ...change }) => {
        if (!previousFormatting) return change;
        const previousInherited =
          ctx.paragraphNumbering(previousFormatting.styleId) ?? previousFormatting.numPrFromStyle;
        const previousList = resolveParagraphList(
          { numPr: previousFormatting.numPr, numPrFromStyle: previousInherited },
          numbering,
        );
        const previousRendering =
          previousList.kind === "marker" ? previousList.rendering : undefined;
        originalList ??= previousRendering;
        const previousAttrs = listLabelAttrsFromRendering(previousRendering, {
          numPr: previousFormatting.numPr,
          numPrFromStyle: previousInherited,
        });
        return {
          ...change,
          previousFormatting: {
            ...previousFormatting,
            ...previousAttrs,
            ...(previousAttrs.numPr === undefined
              ? { numPr: null }
              : { numPr: previousAttrs.numPr }),
            ...(previousAttrs.numPrFromStyle === undefined
              ? { numPrFromStyle: null }
              : { numPrFromStyle: previousAttrs.numPrFromStyle }),
          },
        };
      },
    );
  }
  const { label } = ctx.nextListLabel(counterAttrs);
  const inline = renderParagraphInline(ctx, pkg, para.content, para.paraId);
  const styleId = para.formatting?.styleId;

  const headingLevel = markdownHeadingLevel(ctx, para);
  if (headingLevel !== undefined) {
    ctx.listIndentWidths = [];
    // A numbered heading (`1. Scope`, through its style's `w:numPr` or its
    // own) keeps its number; a bulleted one its glyph, as a heading has no
    // Markdown bullet syntax to borrow.
    if (!inline || (!inline.trim() && label === undefined)) {
      // Whitespace without a visible list label carries no text block.
      return { type: "paragraph", markdown: "" };
    }
    const hashes = "#".repeat(Math.min(MAX_MARKDOWN_HEADING_LEVEL, headingLevel + 1));
    // ATX headings occupy one source line; physical newlines create new blocks.
    const headingInline = inlineBreaksToHtml(inline);
    return {
      markdown: label ? `${hashes} ${label} ${headingInline}` : `${hashes} ${headingInline}`,
      type: "paragraph",
    };
  }

  const visibleList = resolvedRendering ?? originalList;
  if (visibleList && label !== undefined) {
    return renderListItem(ctx, visibleList, label, inline);
  }

  ctx.listIndentWidths = [];

  if (isQuoteStyle(styleId, ctx.builtInStyles)) {
    return {
      markdown: inline
        .split("\n")
        .map((line) => `> ${line}`)
        .join("\n"),
      type: "paragraph",
    };
  }

  return { type: "paragraph", markdown: escapeLeadingBlockMarker(inline) };
}

/** The paragraph's 0-based heading level, as `builtInStyles` classifies it. */
function markdownHeadingLevel(ctx: RenderContext, para: Paragraph): number | undefined {
  return resolveHeadingLevel(
    { outlineLevel: para.formatting?.outlineLevel, styleId: para.formatting?.styleId },
    ctx.builtInStyles,
  );
}

/**
 * A plain paragraph whose visible text begins with markdown block syntax (e.g.
 * `# Not a heading`, `- value`, `1. value`, `> quote`) would be reclassified as
 * a heading/list/blockquote on re-parse, even though Word carries no matching
 * style/list metadata. Escape the leading marker so literal text round-trips.
 */
function escapeLeadingBlockMarker(text: string): string {
  // Escape at the start of every line (`m` flag), not just the paragraph: an
  // inline break (soft break / page break) can put block syntax at the start of
  // a later line. Only horizontal whitespace `[ \t]` leads the marker so the
  // anchor doesn't cross the newline.
  return text
    .replace(/^(?<ws>[ \t]*)(?<marker>[#>=])/gmu, "$<ws>\\$<marker>")
    .replace(/^(?<ws>[ \t]*)(?<marker>[-+*])(?=[ \t]|$)/gmu, "$<ws>\\$<marker>")
    .replace(/^(?<ws>[ \t]*)-(?=[- \t]*$)/gmu, "$<ws>\\-")
    .replace(
      /^(?<ws>[ \t]*)(?<num>\d{1,9})(?<marker>[.)])(?<after>[ \t])/gmu,
      "$<ws>$<num>\\$<marker>$<after>",
    );
}

/**
 * Only emitted Markdown ancestors contribute columns. OOXML levels may skip
 * ancestors or resume after prose; inventing those columns creates code blocks.
 */
function listIndentWidth(ctx: RenderContext, level: number): number {
  let width = 0;
  for (let ancestor = 0; ancestor < level; ancestor++) {
    width += ctx.listIndentWidths[ancestor] ?? 0;
  }
  return width;
}

/**
 * A list item line with the label the page shows (`1.`, `a)`, `i.`), counted
 * in document order across the whole render, list items inside tables and
 * block SDTs included. A bullet is Markdown's `-`.
 *
 * The indent before the marker is not a fixed two spaces: CommonMark keeps a
 * child nested under its parent only when it starts at or past the column
 * where the parent item's content begins, which is the width of the parent's
 * marker plus its one trailing space (`1. ` is 3 columns, `10. ` is 4, `- ` is
 * 2). `ctx.listIndentWidths` is updated with this item's own marker width so
 * a following, deeper item can indent correctly under it.
 */
function renderListItem(
  ctx: RenderContext,
  list: ListRendering,
  label: string,
  inline: string,
): RenderedParagraph {
  // Custom Word labels are paragraph text in CommonMark, not nesting parents.
  const nativeMarker = list.isBullet || /^\d{1,9}[.)]$/u.test(label);
  if (!nativeMarker) {
    ctx.listIndentWidths = [];
    return {
      type: "paragraph",
      markdown: escapeLeadingBlockMarker(`${label} ${inline}`).trimEnd(),
    };
  }
  ctx.listIndentWidths.length = list.level;
  const indent = " ".repeat(listIndentWidth(ctx, list.level));
  const marker = list.isBullet ? "- " : `${label} `;
  ctx.listIndentWidths[list.level] = marker.length;
  const markdown = `${indent}${marker}${escapeLeadingBlockMarker(inline)}`.trimEnd();
  return {
    type: "listItem",
    markdown,
    // Empty list markers cannot interrupt paragraph text; a nested hyphen
    // instead becomes a Setext underline. Separate every empty item.
    separation: inline.trim().length === 0 ? "loose" : "tight",
  };
}
