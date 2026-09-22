/**
 * Render a Word table as markdown. Two output modes, picked automatically per
 * table: GFM for simple tables (no merged cells, no nested tables); inline HTML
 * `<table>` with `colspan`/`rowspan` when the table has `gridSpan`, `vMerge`, or
 * a nested table. Multi-paragraph cells join their paragraphs with `<br>` in
 * both modes. Ported from eigenpal/docx-editor PR #595.
 */

import type {
  BlockContent,
  DocxPackage,
  Hyperlink,
  ParagraphContent,
  Run,
  TrackedRunContent,
  Table,
  TableCell,
  TableRow,
} from "../types/document";
import { decodeOoxmlSymbolCharacter } from "../utils/ooxmlSymbol";
import { escapeTableCell } from "./escape";
import { registerImage } from "./images";
import { getHyperlinkRuns } from "../docx/hyperlinkParser";
import { RELATIONSHIP_TYPES, resolveRelationshipIdOfType } from "../docx/relsParser";
import { pushWarning } from "./internals";
import { renderParagraph } from "./renderParagraph";
import type { RenderContext } from "./types";

/** Render a `Table` as markdown. Picks GFM vs HTML based on cell features. */
export function renderTable(
  ctx: RenderContext,
  pkg: DocxPackage | undefined,
  table: Table,
): string {
  const { rows } = table;
  if (!rows.length) {
    return "";
  }
  if (needsHtmlFallback(rows)) {
    return renderHtmlTable(ctx, pkg, rows, true);
  }
  return renderGfmTable(ctx, pkg, rows, true);
}

function needsHtmlFallback(rows: TableRow[]): boolean {
  for (const row of rows) {
    for (const cell of row.cells) {
      if ((cell.formatting?.gridSpan ?? 1) > 1) {
        return true;
      }
      if (cell.formatting?.vMerge) {
        return true;
      }
      if (containsTable(cell.content)) {
        return true;
      }
    }
  }
  return false;
}

const containsTable = (blocks: readonly BlockContent[]): boolean =>
  blocks.some(
    (block) =>
      block.type === "table" || (block.type === "blockSdt" && containsTable(block.content)),
  );

// ---------------------------------------------------------------------------
// GFM path
// ---------------------------------------------------------------------------

function renderGfmTable(
  ctx: RenderContext,
  pkg: DocxPackage | undefined,
  rows: TableRow[],
  firstRowIsHeader: boolean,
): string {
  const cellTexts = rows.map((row) => renderGfmRow(ctx, pkg, row));
  let maxCols = 0;
  for (const cells of cellTexts) {
    if (cells.length > maxCols) {
      maxCols = cells.length;
    }
  }
  if (!maxCols) {
    return "";
  }

  const padded = cellTexts.map((row) => {
    const filled = row.slice();
    while (filled.length < maxCols) {
      filled.push("");
    }
    return filled;
  });

  const separator = `| ${Array.from({ length: maxCols }, () => "---").join(" | ")} |`;
  const lines: string[] = [];
  if (firstRowIsHeader) {
    const [header, ...body] = padded;
    if (header) {
      lines.push(toRowLine(header));
    }
    lines.push(separator);
    for (const row of body) {
      lines.push(toRowLine(row));
    }
  } else {
    lines.push(`| ${Array.from({ length: maxCols }, () => "").join(" | ")} |`);
    lines.push(separator);
    for (const row of padded) {
      lines.push(toRowLine(row));
    }
  }
  return lines.join("\n");
}

function toRowLine(cells: string[]): string {
  return `| ${cells.join(" | ")} |`;
}

function renderGfmRow(ctx: RenderContext, pkg: DocxPackage | undefined, row: TableRow): string[] {
  const out: string[] = [];
  for (const cell of row.cells) {
    out.push(renderGfmCell(ctx, pkg, cell));
  }
  return out;
}

function renderGfmCell(ctx: RenderContext, pkg: DocxPackage | undefined, cell: TableCell): string {
  const blocks: string[] = [];
  const renderBlocks = (content: readonly BlockContent[]): void => {
    for (const item of content) {
      switch (item.type) {
        case "paragraph": {
          const md = renderParagraph(ctx, pkg, item);
          if (md.trim()) {
            blocks.push(md);
          }
          break;
        }
        case "blockSdt":
          renderBlocks(item.content);
          break;
        case "table":
        case "preservedBlock":
        case "bookmarkStart":
        case "bookmarkEnd":
          break;
        default: {
          const unsupported: never = item;
          return unsupported;
        }
      }
    }
  };
  renderBlocks(cell.content);
  return escapeTableCell(blocks.join("\n"));
}

// ---------------------------------------------------------------------------
// HTML path
// ---------------------------------------------------------------------------

function renderHtmlTable(
  ctx: RenderContext,
  pkg: DocxPackage | undefined,
  rows: TableRow[],
  firstRowIsHeader: boolean,
): string {
  const out: string[] = ["<table>"];
  for (const [rowIdx, row] of rows.entries()) {
    const tag = firstRowIsHeader && rowIdx === 0 ? "th" : "td";
    out.push("  <tr>");
    let gridCol = 0;
    for (const cell of row.cells) {
      const colspan = cell.formatting?.gridSpan ?? 1;
      if (cell.formatting?.vMerge !== "continue") {
        const rowspan = countRowSpan(rows, rowIdx, gridCol, colspan);
        const attrs: string[] = [];
        if (colspan > 1) {
          attrs.push(`colspan="${colspan}"`);
        }
        if (rowspan > 1) {
          attrs.push(`rowspan="${rowspan}"`);
        }
        const attrStr = attrs.length ? ` ${attrs.join(" ")}` : "";
        out.push(`    <${tag}${attrStr}>${renderHtmlCell(ctx, pkg, cell)}</${tag}>`);
      }
      gridCol += colspan;
    }
    out.push("  </tr>");
  }
  out.push("</table>");
  return out.join("\n");
}

/**
 * Count how many rows beyond `rowIdx` have a `vMerge: "continue"` cell aligned
 * to the same grid column as the anchor. Cells are indexed by array position,
 * but vertical merges align on the visual grid column, so a horizontal merge
 * (`gridSpan > 1`) above shifts array indices — we walk by cumulative gridSpan.
 */
function countRowSpan(rows: TableRow[], rowIdx: number, gridCol: number, colspan: number): number {
  let span = 1;
  for (let r = rowIdx + 1; r < rows.length; r++) {
    const row = rows[r];
    if (!row) {
      break;
    }
    const target = cellAtGridColumn(row, gridCol);
    if (
      target &&
      target.formatting?.vMerge === "continue" &&
      (target.formatting.gridSpan ?? 1) === colspan
    ) {
      span += 1;
    } else {
      break;
    }
  }
  return span;
}

function cellAtGridColumn(row: TableRow, gridCol: number): TableCell | undefined {
  let col = 0;
  for (const cell of row.cells) {
    if (col === gridCol) {
      return cell;
    }
    col += cell.formatting?.gridSpan ?? 1;
    if (col > gridCol) {
      return undefined;
    }
  }
  return undefined;
}

function renderHtmlCell(ctx: RenderContext, pkg: DocxPackage | undefined, cell: TableCell): string {
  const parts: string[] = [];
  const renderBlocks = (content: readonly BlockContent[]): void => {
    for (const item of content) {
      switch (item.type) {
        case "paragraph": {
          const inner = renderHtmlInline(ctx, pkg, item.content, item.paraId);
          if (inner) {
            parts.push(inner);
          }
          break;
        }
        case "table": {
          // Nested tables inside an HTML cell stay HTML: GFM is not parsed inside
          // HTML blocks, so a pipe-table here would render as literal text.
          const nested = renderHtmlTable(ctx, pkg, item.rows, true);
          if (nested) {
            parts.push(nested);
          }
          break;
        }
        case "blockSdt":
          renderBlocks(item.content);
          break;
        case "preservedBlock":
        case "bookmarkStart":
        case "bookmarkEnd":
          break;
        default: {
          const unsupported: never = item;
          return unsupported;
        }
      }
    }
  };
  renderBlocks(cell.content);
  return parts.join("<br>");
}

// ---------------------------------------------------------------------------
// HTML inline rendering for table cells. Markdown is not parsed inside HTML
// blocks, so cells inside an HTML `<table>` emit HTML tags for marks and links.
// ---------------------------------------------------------------------------

function escapeHtml(text: string): string {
  return text
    .replace(/&/gu, "&amp;")
    .replace(/</gu, "&lt;")
    .replace(/>/gu, "&gt;")
    .replace(/"/gu, "&quot;");
}

function renderHtmlInline(
  ctx: RenderContext,
  pkg: DocxPackage | undefined,
  content: readonly ParagraphContent[],
  paraId: string | undefined,
): string {
  let out = "";
  for (const item of content) {
    switch (item.type) {
      case "run":
        out += renderHtmlRun(ctx, pkg, item, paraId);
        break;
      case "hyperlink":
        out += renderHtmlHyperlink(ctx, pkg, item, paraId);
        break;
      case "insertion":
      case "moveTo": {
        const inner = renderHtmlChildren(ctx, pkg, item.content, paraId);
        out += ctx.opts.trackedChanges === "annotate" ? `<ins>${inner}</ins>` : inner;
        break;
      }
      case "deletion":
      case "moveFrom":
        if (ctx.opts.trackedChanges === "annotate") {
          out += `<del>${renderHtmlChildren(ctx, pkg, item.content, paraId)}</del>`;
        }
        break;
      case "simpleField":
      case "complexField": {
        const runs = item.type === "simpleField" ? item.content : item.fieldResult;
        out += renderHtmlChildren(ctx, pkg, runs, paraId);
        break;
      }
      // A content control states what its text is bound to and a bidirectional
      // wrapper states how it is laid out; an HTML cell carries neither, so
      // both are read through to the text itself.
      case "inlineSdt":
      case "inlineWrapper":
        out += renderHtmlInline(ctx, pkg, item.content, paraId);
        break;
      case "mathEquation":
        // Markdown can't carry OMML; emit the plain-text fallback when present.
        if (item.plainText) {
          out += escapeHtml(item.plainText);
        }
        break;
      default:
        // Range markers carry no visible payload inside table cells.
        break;
    }
  }
  return out;
}

function renderHtmlChildren(
  ctx: RenderContext,
  pkg: DocxPackage | undefined,
  children: readonly TrackedRunContent[],
  paraId: string | undefined,
): string {
  return children
    .map((child): string => {
      switch (child.type) {
        case "run":
          return renderHtmlRun(ctx, pkg, child, paraId);
        case "hyperlink":
          return renderHtmlHyperlink(ctx, pkg, child, paraId);
        case "mathEquation":
          return child.plainText ? escapeHtml(child.plainText) : "";
        // A transparent wrapper carries the revision's text; reading through
        // it is the only way that text reaches the cell.
        case "inlineWrapper":
        case "inlineSdt":
        case "simpleField":
        case "complexField":
        case "insertion":
        case "deletion":
        case "moveFrom":
        case "moveTo":
          return renderHtmlInline(ctx, pkg, [child], paraId);
        // A range boundary carries no text, and neither does markup folio kept
        // as bytes: a capture is not read, so it has no words to render.
        case "bookmarkStart":
        case "bookmarkEnd":
        case "moveFromRangeStart":
        case "moveFromRangeEnd":
        case "moveToRangeStart":
        case "moveToRangeEnd":
        case "preservedInline":
          return "";
        default: {
          const unrendered: never = child;
          return unrendered;
        }
      }
    })
    .join("");
}

function renderHtmlRun(
  ctx: RenderContext,
  pkg: DocxPackage | undefined,
  run: Run,
  paraId: string | undefined,
): string {
  // Hidden text (`w:vanish`) is suppressed in Word's normal view; drop it.
  if (run.formatting?.hidden) {
    return "";
  }
  let text = "";
  for (const item of run.content) {
    switch (item.type) {
      case "text":
        text += escapeHtml(item.text);
        break;
      case "tab":
        text += "&emsp;";
        break;
      case "break":
        text += "<br>";
        break;
      case "symbol":
        text += escapeHtml(decodeOoxmlSymbolCharacter(item.char) ?? item.char);
        break;
      case "noBreakHyphen":
        text += "&#8209;";
        break;
      case "preservedXml":
        text += escapeHtml(item.text);
        break;
      case "softHyphen":
        break;
      case "drawing": {
        const ref = resolveRelationshipIdOfType(
          pkg?.relationships,
          item.image.rId,
          RELATIONSHIP_TYPES.image,
        );
        const media =
          ref.status === "resolved" ? pkg?.media?.get(ref.relationship.target) : undefined;
        if (media) {
          const reg = registerImage(ctx, media, item.image, paraId);
          const alt = reg.alt ? escapeHtml(reg.alt) : "";
          text += `<img src="${escapeHtml(reg.virtualPath)}" alt="${alt}">`;
          break;
        }
        if (item.image.src) {
          const alt = escapeHtml(item.image.alt ?? item.image.title ?? item.image.filename ?? "");
          text += `<img src="${escapeHtml(item.image.src)}" alt="${alt}">`;
          break;
        }
        pushWarning(ctx, `image rId=${item.image.rId ?? "(absent)"} not resolvable`);
        break;
      }
      case "footnoteRef":
      case "endnoteRef": {
        if (ctx.opts.footnotes === "strip") {
          break;
        }
        const markerNumber = ctx.footnoteRefs.length + 1;
        ctx.footnoteRefs.push({
          refId: item.id,
          markerNumber,
          kind: item.type === "endnoteRef" ? "endnote" : "footnote",
        });
        text += `<sup>[${markerNumber}]</sup>`;
        break;
      }
      default:
        break;
    }
  }
  if (!text) {
    return "";
  }
  const f = run.formatting;
  if (!f) {
    return text;
  }
  if (f.bold) {
    text = `<strong>${text}</strong>`;
  }
  if (f.italic) {
    text = `<em>${text}</em>`;
  }
  if (f.strike) {
    text = `<s>${text}</s>`;
  }
  if (f.underline?.style && f.underline.style !== "none") {
    text = `<u>${text}</u>`;
  }
  return text;
}

function renderHtmlHyperlink(
  ctx: RenderContext,
  pkg: DocxPackage | undefined,
  link: Hyperlink,
  paraId: string | undefined,
): string {
  // Hyperlink.children may include BookmarkStart/End markers; only runs
  // contribute visible content, and a transparent wrapper is read through to
  // the runs it holds.
  const inner = getHyperlinkRuns(link)
    .map((run) => renderHtmlRun(ctx, pkg, run, paraId))
    .join("");
  if (!inner) {
    return "";
  }
  const href = link.href ?? (link.anchor ? `#${link.anchor}` : "");
  if (!href) {
    return inner;
  }
  return `<a href="${escapeHtml(href)}">${inner}</a>`;
}
