/** Test-only ops/editor oracle. Disagreements are data, never successful comparisons. */
import type { Document } from "@stll/docx-core/model";
import {
  applyDocumentOp,
  normalizeForOps,
  DOCUMENT_OP_TYPES,
  OP_STORIES,
  type RevisionDecision,
  validateOpsDocument,
} from "@stll/docx-core/ops";
import { Result } from "better-result";
import { EditorState } from "prosemirror-state";

import type { PackageDifferences } from "../../../../../scripts/lib/corpus-invariants/model-equality";
import { reviewDifferences } from "../../../../../test/reviewDifferences";
import { storyRevisionIds } from "../../../../../test/reviewProjection";
import { repackDocx } from "../../docx/rezip";
import { parseDocx } from "../../docx/parser";
import { resolveAllChangesInHeadlessState } from "../../prosemirror/commands/comments";
import { updateDocumentContent } from "../../prosemirror/conversion/fromProseDoc";
import { createDocumentStylesPlugin } from "../../prosemirror/plugins/documentStyles";
import { createDocumentNumberingPlugin } from "../../prosemirror/plugins/documentNumbering";
import { toProseDoc } from "../../prosemirror/conversion/toProseDoc";

export { storyRevisionIds, reviewDifferences };

export const editorProjectionRoundTrip = (document: Document): Document => {
  const projected = toProseDoc(document);
  return updateDocumentContent(document, projected);
};

export const editorRoundTrip = async (document: Document): Promise<Document> => {
  const rebuilt = editorProjectionRoundTrip(document);
  return normalizeForOps(
    await parseDocx(await repackDocx(rebuilt, { updateModifiedDate: false }), {
      preloadFonts: false,
      detectVariables: false,
    }),
  );
};

export const resolveInEditor = (document: Document, decision: RevisionDecision): Document => {
  const doc = toProseDoc(document);
  const state = EditorState.create({
    schema: doc.type.schema,
    doc,
    plugins: [
      createDocumentStylesPlugin(document.package.styles),
      createDocumentNumberingPlugin(document.package.numbering),
    ],
  });
  const resolved = resolveAllChangesInHeadlessState(state, decision);
  return updateDocumentContent(document, resolved.doc);
};

export type ReviewOracleOutcome =
  | { type: "no-revisions" }
  | { type: "invalid"; reason: string }
  | { type: "refused"; decision: RevisionDecision; reason: string }
  | { type: "editor-failed"; decision: RevisionDecision; errorType: string }
  | { type: "match"; decision: RevisionDecision }
  | { type: "disagreement"; decision: RevisionDecision; differences: PackageDifferences };

export const compareReviewResolution = (
  document: Document,
  decision: RevisionDecision,
): ReviewOracleOutcome => {
  const revisionIds = storyRevisionIds(document);
  if (revisionIds.length === 0) return { type: "no-revisions" };
  const contract = validateOpsDocument(document);
  if (contract.isErr()) return { type: "invalid", reason: contract.error.reason };
  const model = applyDocumentOp(document, {
    type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
    story: OP_STORIES.MAIN,
    revisionIds,
    decision,
  });
  if (model.isErr()) return { type: "refused", decision, reason: model.error.reason };
  const editor = Result.try(() => resolveInEditor(document, decision));
  if (editor.isErr()) {
    return {
      type: "editor-failed",
      decision,
      errorType: editor.error instanceof Error ? editor.error.name : "UnknownError",
    };
  }
  const differences = reviewDifferences(model.value.document, editor.value);
  if (differences.messages.length > 0 || differences.omitted > 0) {
    return { type: "disagreement", decision, differences };
  }
  return { type: "match", decision };
};
