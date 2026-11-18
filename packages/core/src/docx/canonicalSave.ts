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
import type {
  CanonicalSaveSnapshot,
  CanonicalSaveSnapshotBaseline,
} from "../types/canonicalSave";
import type { HeaderFooter } from "../types/document";
import {
  FOLIO_DOCX_SERIALIZATION_MODE,
  type FolioDocxSerializationPath,
  type FolioGetDocxOptions,
} from "../types/docxSerialization";
import { repackWithCanonicalStoryRemovals } from "./canonicalStoryRepack";
import {
  copyHeaderFooterBaselineHandles,
  getDocumentSourceBaseline,
} from "./headerFooterVerbatim";
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

/** The adapters' error channel for one fidelity fallback. */
export const canonicalSaveDiagnosticError = (
  diagnostic: SaveDiagnostic,
): CanonicalSaveDiagnosticError =>
  new CanonicalSaveDiagnosticError({
    gap: CANONICAL_GAP.save,
    diagnostic,
    message: `Canonical save used a fidelity fallback: ${diagnostic.type} (${diagnostic.part}).`,
  });

const CanonicalSaveRepackRefusedErrorBase: TaggedErrorClass<"CanonicalSaveRepackRefusedError"> =
  TaggedError("CanonicalSaveRepackRefusedError");

/** A save allowed only to patch the loaded package found that it needs a repack. */
export class CanonicalSaveRepackRefusedError extends CanonicalSaveRepackRefusedErrorBase<{
  message: string;
}> {}

type SerializeCanonicalSaveOptions = {
  snapshot: CanonicalSaveSnapshot;
  /** `"refuse"`: throw {@link CanonicalSaveRepackRefusedError} rather than build a repack. */
  repack?: "allow" | "refuse" | undefined;
  options?: FolioGetDocxOptions | undefined;
  featureFlags?: FolioSelectiveSaveFlags | undefined;
};

/** Model and change signals come from one committed snapshot, never from PM. */
/**
 * `document` saved against `baseline`: its buffer, its body capture, and the
 * verbatim source of each header and footer, so an edit since the baseline's
 * save is judged against what that save wrote.
 */
const withSaveBaseline = (
  document: Document,
  { buffer, document: baseline }: CanonicalSaveSnapshotBaseline,
): Document => {
  const rebaseParts = (
    parts: Map<string, HeaderFooter> | undefined,
    sources: Map<string, HeaderFooter> | undefined,
  ) =>
    parts === undefined
      ? undefined
      : new Map(
          [...parts].map(([rId, part]) => {
            const rebased: HeaderFooter = { ...part };
            delete rebased.verbatimXml;
            delete rebased.verbatimFingerprint;
            const source = sources?.get(rId);
            if (source?.verbatimXml !== undefined)
              rebased.verbatimXml = source.verbatimXml;
            if (source?.verbatimFingerprint !== undefined)
              rebased.verbatimFingerprint = source.verbatimFingerprint;
            return [rId, rebased];
          }),
        );
  const headers = rebaseParts(
    document.package.headers,
    baseline.package.headers,
  );
  const footers = rebaseParts(
    document.package.footers,
    baseline.package.footers,
  );
  const rebased: Document = {
    ...document,
    originalBuffer: buffer,
    package: {
      ...document.package,
      document: { ...document.package.document },
      ...(headers === undefined ? {} : { headers }),
      ...(footers === undefined ? {} : { footers }),
    },
  };
  copyHeaderFooterBaselineHandles(rebased, baseline);
  return rebased;
};

export const serializeCanonicalSave = async ({
  snapshot,
  repack: repackPolicy = "allow",
  options,
  featureFlags,
}: SerializeCanonicalSaveOptions): Promise<{
  buffer: ArrayBuffer;
  /** Whether the buffer is the selective patch of the baseline or a repack. */
  path: FolioDocxSerializationPath;
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
  // The package the session last saved, when it has one, is what this save
  // patches; the result still carries the session's own document.
  const written =
    snapshot.saveBaseline === undefined
      ? document
      : withSaveBaseline(document, snapshot.saveBaseline);
  const baseline = written.originalBuffer;
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
    selectiveBuffer = await attemptSelectiveSave(written, baseline, {
      bodyAuthority: "canonical",
      changedParaIds: new Set(
        snapshot.changedBlockIds.filter((id) => !heldBack.has(idKey(id))),
      ),
      structuralChange: snapshot.structure === "changed",
      hasUntrackedChanges: false,
      maxBytes: flags.selectiveSaveMaxBytes,
      onDiagnostic,
      onSelectiveDecline: options?.onSelectiveDecline,
    });
    if (useSelective && !selectiveBuffer)
      onDiagnostic({ type: "selectiveSaveRefused", part: "word/document.xml" });
  }
  if (repackPolicy === "refuse" && !(useSelective && selectiveBuffer))
    throw new CanonicalSaveRepackRefusedError({
      message: "The save needs a repack of the package, which was refused.",
    });
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
    if (baseline && getDocumentSourceBaseline(written).type === "missing")
      onDiagnostic({
        type: "sourceReplayUnavailable",
        part: "word/document.xml",
      });
    return baseline
      ? repackWithCanonicalStoryRemovals({
          document: written,
          repack: () =>
            repackDocx(written, {
              onDiagnostic,
              bodyAuthority: "canonical",
              changedNoteParaIds,
            }),
        })
      : createDocx(written, { onDiagnostic, bodyAuthority: "canonical" });
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
  const path = fullBuffer === buffer ? "full-repack" : "selective";
  options?.onSerializationPath?.(path);
  return {
    buffer,
    path,
    document,
    version: snapshot.version,
    diagnostics,
    tripwireResult,
  };
};
