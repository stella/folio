import {
  canonicalResourceReplacementOf,
  CanonicalResourceSaveRefusalError,
} from "./canonicalResourceSave";
import { Result, TaggedError } from "better-result";
import {
  FOLIO_DOCX_SERIALIZATION_MODE,
  type FolioGetDocxOptions,
} from "../types/docxSerialization";
import { CANONICAL_GAP } from "../types/canonicalCapabilities";
import type { CanonicalSaveSnapshot } from "../types/canonicalSave";
import { withoutUnreferencedNotes } from "../prosemirror/noteReferenceReview";
import { repackWithCanonicalStoryRemovals } from "./canonicalStoryRepack";
import { getDocumentSourceBaseline } from "./headerFooterVerbatim";
import type { SaveDiagnostic } from "./saveDiagnostics";
import type { FolioSelectiveSaveFlags } from "./selectiveSaveFlags";
import { resolveSelectiveSaveFlags } from "./selectiveSaveFlags";

/** Forward a fidelity fallback through the adapters' existing error channel. */
export class CanonicalSaveDiagnosticError extends TaggedError("CanonicalSaveDiagnosticError")<{
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
}: SerializeCanonicalSaveOptions) => {
  const document = withoutUnreferencedNotes(snapshot.document);
  // Package-resource operations conservatively retain structural save work through undo.
  const replacementPart =
    snapshot.structure === "changed" ? canonicalResourceReplacementOf(document) : undefined;
  if (replacementPart) {
    const diagnostic = {
      type: "canonicalResourceReplacement",
      gap: CANONICAL_GAP.resourceReplacement,
      part: replacementPart,
    } as const satisfies SaveDiagnostic;
    options?.onDiagnostic?.(diagnostic);
    throw new CanonicalResourceSaveRefusalError({
      message: "Canonical save cannot preserve this package resource replacement.",
      gap: CANONICAL_GAP.resourceReplacement,
      diagnostic,
    });
  }
  const baseline = document.originalBuffer;
  const flags = resolveSelectiveSaveFlags(featureFlags);
  const diagnostics: SaveDiagnostic[] = [];
  const onDiagnostic = (diagnostic: SaveDiagnostic) => {
    if (
      !diagnostics.some(
        (existing) => existing.type === diagnostic.type && existing.part === diagnostic.part,
      )
    )
      diagnostics.push(diagnostic);
  };
  const { repackDocx, createDocx } = await import("./rezip");
  const useSelective = flags.selectiveSave && options?.mode !== FOLIO_DOCX_SERIALIZATION_MODE.full;
  let selectiveBuffer: ArrayBuffer | null = null;
  if (baseline && (useSelective || flags.selectiveSaveTripwire)) {
    const { attemptSelectiveSave } = await import("./selectiveSave");
    selectiveBuffer = await attemptSelectiveSave(document, baseline, {
      bodyAuthority: "canonical",
      changedParaIds: new Set(snapshot.changedBlockIds),
      structuralChange: snapshot.structure === "changed",
      hasUntrackedChanges: false,
      maxBytes: flags.selectiveSaveMaxBytes,
      onDiagnostic,
    });
    if (useSelective && !selectiveBuffer)
      onDiagnostic({ type: "selectiveSaveRefused", part: "word/document.xml" });
  }
  const repack = () => {
    if (baseline && getDocumentSourceBaseline(document).type === "missing")
      onDiagnostic({ type: "sourceReplayUnavailable", part: "word/document.xml" });
    return baseline
      ? repackWithCanonicalStoryRemovals({
          document,
          repack: () => repackDocx(document, { onDiagnostic, bodyAuthority: "canonical" }),
        })
      : createDocx(document, { onDiagnostic, bodyAuthority: "canonical" });
  };
  let buffer = useSelective ? selectiveBuffer : null;
  let fullBuffer: ArrayBuffer | null = null;
  if (!buffer) {
    fullBuffer = await repack();
    buffer = fullBuffer;
  } else if (flags.selectiveSaveTripwire) {
    const full = await Result.tryPromise({ try: repack, catch: (error: unknown) => error });
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
  return { buffer, document, version: snapshot.version, diagnostics, tripwireResult };
};
