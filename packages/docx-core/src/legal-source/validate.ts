import type { LegalDraft, LegalDraftBlock, LegalDraftDiagnostic } from "./types";

// `**whole paragraph**` or `__whole paragraph__`: emphasis wrapping every
// character of a body string.
const WHOLE_TEXT_EMPHASIS = /^(?:\*\*[^*]+\*\*|__[^_]+__)$/u;

/** Body strings of a block that a reader would expect to be prose, not a heading. */
const bodyStrings = (block: LegalDraftBlock): string[] => {
  switch (block.type) {
    case "paragraph":
    case "recital":
    case "clause":
    case "schedule":
      return block.paragraphs;
    case "list":
      return block.items;
    case "table":
      return block.table.rows.flat();
    case "title":
    case "signatures":
    case "pageBreak":
      return [];
    default:
      block satisfies never;
      return [];
  }
};

export const validateLegalDraft = (draft: LegalDraft): LegalDraftDiagnostic[] => {
  const diagnostics: LegalDraftDiagnostic[] = [];

  // A body paragraph bold from end to end usually carries chat prose over
  // into the draft; in a two-column bilingual table it turns a whole column
  // into a heading. The compiler renders it as written and reports it, so
  // the author can decide.
  for (const block of draft.blocks) {
    if (bodyStrings(block).some((text) => WHOLE_TEXT_EMPHASIS.test(text.trim()))) {
      diagnostics.push({
        code: "whole-paragraph-emphasis",
        message:
          "A body paragraph, list item, or table cell is bold from end to end; use a clause heading for headings and keep bold for short labels.",
        severity: "warning",
      });
      break;
    }
  }

  if (!draft.meta.title?.trim()) {
    diagnostics.push({
      code: "missing-title",
      message: "The draft must have a title.",
      severity: "error",
    });
  }

  let hasTopLevelClause = false;
  for (const block of draft.blocks) {
    if (block.type === "clause") {
      if (block.level === 1) {
        hasTopLevelClause = true;
      }
      if (block.level > 1 && !hasTopLevelClause) {
        diagnostics.push({
          code: "subclause-before-clause",
          message: "A subclause cannot appear before the first top-level clause.",
          severity: "error",
        });
      }
      if (!block.heading.trim()) {
        diagnostics.push({
          code: "empty-clause-heading",
          message: "Clause headings cannot be empty.",
          severity: "error",
        });
      }
    }

    if (block.type === "table") {
      if (block.table.headers.length === 0) {
        diagnostics.push({
          code: "empty-table",
          message: "Tables must include at least one header.",
          severity: "error",
        });
      }
      for (const row of block.table.rows) {
        if (row.length !== block.table.headers.length) {
          diagnostics.push({
            code: "ragged-table",
            message: "Every table row must match the header width.",
            severity: "error",
          });
        }
      }
    }

    if (block.type === "signatures" && block.parties.length === 0) {
      diagnostics.push({
        code: "empty-signatures",
        message: "The signatures block must include at least one party.",
        severity: "warning",
      });
    }
  }

  return diagnostics;
};
