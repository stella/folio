import type { FolioReviewChange } from "@stll/folio-core/server";

import type { FolioAgentChange } from "../types";

/**
 * Map a core {@link FolioReviewChange} — from `FolioDocxReviewer.getChanges`
 * (headless bridge, `bridges/reviewer.ts`) or `DocxEditorRef.getTrackedChanges`
 * (live-editor bridge, `bridges/editor-ref.ts`) — to the tool-facing
 * {@link FolioAgentChange} shape. Both bridges read the same underlying
 * tracked-change record (`getTrackedChangesFromDoc` in
 * `@stll/folio-core/ai-edits`), so they share this one mapping instead of
 * drifting into two subtly different shapes.
 */
export const toAgentChange = (change: FolioReviewChange): FolioAgentChange => ({
  id: String(change.id),
  type: change.type,
  author: change.author,
  text: change.text,
  blockId: change.blockId,
});
