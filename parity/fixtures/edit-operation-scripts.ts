/** Synthetic saved-edit scenarios. Each script starts from its own package. */

export const EDIT_OPERATION_SCRIPTS = [
  {
    id: "delete-final-paragraph",
    seed: "edit-final-paragraph-seed.docx",
    action: "deleteBlock",
    anchorText: "Target final paragraph Omega 47.",
    mode: "direct",
  },
  {
    id: "accept-deleted-paragraph-boundary",
    seed: "edit-paragraph-boundary-seed.docx",
    action: "mergeBlockWithNext",
    anchorText: "Paragraph before boundary Cedar 19.",
    mode: "tracked-changes",
    resolution: "acceptAll",
  },
  {
    id: "insert-continuing-numbering",
    seed: "edit-numbering-seed.docx",
    action: "insertAfterBlock",
    anchorText: "Primary numbered item 23.",
    text: "Inserted primary item 29.",
    mode: "direct",
  },
  {
    id: "delete-row-next-to-merge",
    seed: "edit-merged-table-seed.docx",
    action: "deleteTableRow",
    anchorText: "Beta 52",
    mode: "direct",
  },
  {
    id: "reply-on-comment-range",
    seed: "edit-comment-range-seed.docx",
    action: "replyToComment",
    anchorText: "Review span 18 begins here.",
    text: "Synthetic reply 27.",
  },
  {
    id: "edit-next-to-notes-fields-sections",
    seed: "edit-notes-fields-sections-seed.docx",
    action: "replaceInBlock",
    anchorText: "Second section 61.",
    find: "61",
    replace: "62",
    mode: "direct",
  },
] as const;

export type EditOperationScript = (typeof EDIT_OPERATION_SCRIPTS)[number];
