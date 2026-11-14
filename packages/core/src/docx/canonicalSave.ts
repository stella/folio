import { Result, TaggedError, type TaggedErrorClass } from "better-result";

import type { Document } from "@stll/docx-core";
import {
  idKey,
  packageParagraphIds,
  storyParagraphs,
  withoutPendingSuggestions,
} from "@stll/docx-core/ops";

import { withoutUnreferencedNotes } from "../prosemirror/noteReferenceReview";
import { CANONICAL_GAP } from "../types/canonicalCapabilities";
import type { CanonicalSaveSnapshot } from "../types/canonicalSave";
import {
  FOLIO_DOCX_SERIALIZATION_MODE,
  type FolioGetDocxOptions,
} from "../types/docxSerialization";
import { repackWithCanonicalStoryRemovals } from "./canonicalStoryRepack";
import { getDocumentSourceBaseline } from "./headerFooterVerbatim";
import type { SaveDiagnostic } from "./saveDiagnostics";
import type { FolioSelectiveSaveFlags } from "./selectiveSaveFlags";
import { resolveSelectiveSaveFlags } from "./selectiveSaveFlags";
import type { TripwireResult } from "./selectiveSaveTripwire";
import { withoutUnusedSessionLists } from "./sessionListPruning";

const CanonicalSaveDiagnosticErrorBase: TaggedErrorClass<"CanonicalSaveDiagnosticError"> =
  TaggedError("CanonicalSaveDiagnosticError");

/** Forward a fidelity fallback through the adapters' existing error channel. */
export class CanonicalSaveDiagnosticError extends CanonicalSaveDiagnosticErrorBase<{
  message: string;
  gap: typeof CANONICAL_GAP.save;
  diagnostic: SaveDiagnostic;
}> {}

type SerializeCanonicalSaveOptions = {
  snapshot: CanonicalSaveSnapshot;
  options?: FolioGetDocxOptions | undefined;
  featureFlags?: FolioSelectiveSaveFlags | undefined;
};

/** Model and change signals come from one committed snapshot, never from PM. */
export const serializeCanonicalSave = async ({
  snapshot,
  options,
  featureFlags,
}: SerializeCanonicalSaveOptions): Promise<{
  buffer: ArrayBuffer;
  /** Whether the buffer is the selective patch of the baseline or a repack. */
  path: "selective" | "full-repack";
  document: Document;
  version: number;
  diagnostics: SaveDiagnostic[];
  tripwireResult: TripwireResult | null;
}> => {
  // A pending suggestion is the host's proposal, not a change in the file:
  // the file is written as if every one were rejected. A suggestion the
  // engine cannot reject refuses the save rather than reach the file as an
  // ordinary tracked change.
  const withoutSuggestions = withoutPendingSuggestions(snapshot.document);
  if (withoutSuggestions.isErr()) throw withoutSuggestions.error;
  const document = withoutUnusedSessionLists(
    withoutUnreferencedNotes(
      withoutSuggestions.value,
      snapshot.heldUnreferencedNotes,
    ),
  );
  // A paragraph a pending suggestion adds is not in the file, so it is no
  // change to write: rejecting the suggestion only ever removes paragraphs
  // the suggestion itself added.
  const heldBack = new Set<string>();
  if (withoutSuggestions.value !== snapshot.document) {
    const kept = new Set(
      packageParagraphIds(withoutSuggestions.value.package).map(idKey),
    );
    for (const id of packageParagraphIds(snapshot.document.package))
      if (!kept.has(idKey(id))) heldBack.add(idKey(id));
  }
  const baseline = document.originalBuffer;
  const flags = resolveSelectiveSaveFlags(featureFlags);
  const diagnostics: SaveDiagnostic[] = [];
  const onDiagnostic = (diagnostic: SaveDiagnostic) => {
    if (
      !diagnostics.some(
        (existing) =>
          existing.type === diagnostic.type &&
          existing.part === diagnostic.part,
      )
    )
      diagnostics.push(diagnostic);
  };
  const { repackDocx, createDocx } = await import("./rezip");
  const useSelective =
    flags.selectiveSave && options?.mode !== FOLIO_DOCX_SERIALIZATION_MODE.full;
  let selectiveBuffer: ArrayBuffer | null = null;
  if (baseline && (useSelective || flags.selectiveSaveTripwire)) {
    const { attemptSelectiveSave } = await import("./selectiveSave");
    selectiveBuffer = await attemptSelectiveSave(document, baseline, {
      bodyAuthority: "canonical",
      changedParaIds: new Set(
        snapshot.changedBlockIds.filter((id) => !heldBack.has(idKey(id))),
      ),
      structuralChange: snapshot.structure === "changed",
      hasUntrackedChanges: false,
      maxBytes: flags.selectiveSaveMaxBytes,
      onDiagnostic,
    });
    if (useSelective && !selectiveBuffer)
      onDiagnostic({ type: "selectiveSaveRefused", part: "word/document.xml" });
  }
  // Which note paragraphs changed is the session's journal, as for the body,
  // never a diff of the notes' re-serialization: a note the session never
  // touched (one the package held without a reference, say) is written as the
  // package held it, even where the model cannot spell its source markup.
  // Every paragraph of a note counts, at any depth: one in a table cell or a
  // content control is as much the note's as one at its top level.
  const touched = new Set(snapshot.changedBlockIds.map(idKey));
  const changedNoteParaIds = new Set<string>();
  for (const note of [
    ...(document.package.footnotes ?? []),
    ...(document.package.endnotes ?? []),
  ])
    for (const { paragraph } of storyParagraphs({ content: note.content }))
      if (
        paragraph.paraId !== undefined &&
        touched.has(idKey(paragraph.paraId))
      )
        changedNoteParaIds.add(paragraph.paraId);
  const repack = () => {
    if (baseline && getDocumentSourceBaseline(document).type === "missing")
      onDiagnostic({
        type: "sourceReplayUnavailable",
        part: "word/document.xml",
      });
    return baseline
      ? repackWithCanonicalStoryRemovals({
          document,
          repack: () =>
            repackDocx(document, {
              onDiagnostic,
              bodyAuthority: "canonical",
              changedNoteParaIds,
            }),
        })
      : createDocx(document, { onDiagnostic, bodyAuthority: "canonical" });
  };
  let buffer = useSelective ? selectiveBuffer : null;
  let fullBuffer: ArrayBuffer | null = null;
  if (!buffer) {
    fullBuffer = await repack();
    buffer = fullBuffer;
  } else if (flags.selectiveSaveTripwire) {
    const full = await Result.tryPromise({
      try: repack,
      catch: (error: unknown) => error,
    });
    if (full.isOk()) fullBuffer = full.value;
  }
  let tripwireResult = null;
  if (flags.selectiveSaveTripwire && fullBuffer) {
    const { compareSelectiveVsFull } = await import("./selectiveSaveTripwire");
    const capturedFullBuffer = fullBuffer;
    const compared = await Result.tryPromise({
      try: () => compareSelectiveVsFull(selectiveBuffer, capturedFullBuffer),
      catch: (error: unknown) => error,
    });
    if (compared.isOk()) tripwireResult = compared.value;
  }
  for (const diagnostic of diagnostics) options?.onDiagnostic?.(diagnostic);
  return {
    buffer,
    path: fullBuffer === buffer ? "full-repack" : "selective",
    document,
    version: snapshot.version,
    diagnostics,
    tripwireResult,
  };
};
