import { Result, TaggedError, type TaggedErrorClass } from "better-result";

import type { Document } from "@stll/docx-core";
import {
  idKey,
  packageParagraphIds,
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
import { canonicalJson } from "../utils/canonicalJson";
import { repackWithCanonicalStoryRemovals } from "./canonicalStoryRepack";
import {
  copyHeaderFooterBaselineHandles,
  getDocumentSourceBaseline,
} from "./headerFooterVerbatim";
import { withNormalizedSourceParts } from "./normalizedSourceParts";
import type { SaveDiagnostic } from "./saveDiagnostics";
import type { FolioSelectiveSaveFlags } from "./selectiveSaveFlags";
import { resolveSelectiveSaveFlags } from "./selectiveSaveFlags";
import type { TripwireResult } from "./selectiveSaveTripwire";
import { withoutUnusedSessionLists } from "./sessionListPruning";
import { withoutUnreferencedHeaderFooters } from "./unreferencedHeaderFooters";

const CanonicalSaveDiagnosticErrorBase: TaggedErrorClass<"CanonicalSaveDiagnosticError"> =
  TaggedError("CanonicalSaveDiagnosticError");

/** Forward a fidelity fallback through the adapters' existing error channel. */
export class CanonicalSaveDiagnosticError extends CanonicalSaveDiagnosticErrorBase<{
  message: string;
  gap: typeof CANONICAL_GAP.saveSerialization;
  diagnostic: SaveDiagnostic;
}> {}

/** The adapters' error channel for one fidelity fallback. */
export const canonicalSaveDiagnosticError = (
  diagnostic: SaveDiagnostic,
): CanonicalSaveDiagnosticError =>
  new CanonicalSaveDiagnosticError({
    gap: CANONICAL_GAP.saveSerialization,
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

/** Each paragraph record of a package by id key, nested ones included. */
const paragraphRecordsById = (
  document: Document,
): Map<string, { paraId: string; record: object }> => {
  const out = new Map<string, { paraId: string; record: object }>();
  const visit = (value: unknown): void => {
    if (typeof value !== "object" || value === null) return;
    if (Array.isArray(value)) {
      for (const item of value) visit(item);
      return;
    }
    if (value instanceof Map) {
      for (const item of value.values()) visit(item);
      return;
    }
    if ("type" in value && value.type === "paragraph" && "paraId" in value) {
      const { paraId } = value;
      if (typeof paraId === "string")
        out.set(idKey(paraId), { paraId, record: value });
    }
    for (const field of Object.values(value)) visit(field);
  };
  // The body's sections are a derived view of its content.
  visit({
    ...document.package,
    document: { ...document.package.document, sections: undefined },
  });
  return out;
};

type ParagraphsRewrittenOptions = { before: Document; after: Document };

/**
 * The ids of the paragraphs `after` holds that differ from the paragraph of
 * the same id in `before`. Operations never change a record they are given,
 * so an unchanged paragraph is the same record and costs no comparison.
 */
const paragraphsRewritten = ({
  before,
  after,
}: ParagraphsRewrittenOptions): string[] => {
  const previous = paragraphRecordsById(before);
  const out: string[] = [];
  for (const [key, { paraId, record }] of paragraphRecordsById(after)) {
    const prior = previous.get(key)?.record;
    if (prior === record) continue;
    if (prior !== undefined && canonicalJson(prior) === canonicalJson(record))
      continue;
    out.push(paraId);
  }
  return out;
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
    withoutUnreferencedHeaderFooters(
      withoutUnreferencedNotes(
        withoutSuggestions.value,
        snapshot.heldUnreferencedNotes,
      ),
      snapshot.referencedHeaderFootersAtOpen,
    ),
  );
  // A paragraph a pending suggestion adds is not in the file, so it is no
  // change to write: rejecting the suggestion only ever removes paragraphs
  // the suggestion itself added. Rejecting one can still change a paragraph
  // the session never touched: an accepted deleted break may have run
  // another paragraph's words into a paragraph a suggestion added, and the
  // rejection runs them on into the paragraph that stays. That paragraph is
  // a change to write.
  const heldBack = new Set<string>();
  if (withoutSuggestions.value !== snapshot.document) {
    const kept = new Set(
      packageParagraphIds(withoutSuggestions.value.package).map(idKey),
    );
    for (const id of packageParagraphIds(snapshot.document.package))
      if (!kept.has(idKey(id))) heldBack.add(idKey(id));
  }
  const rejectionRewrote =
    withoutSuggestions.value === snapshot.document
      ? []
      : paragraphsRewritten({
          before: snapshot.document,
          after: withoutSuggestions.value,
        });
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
  const { repackDocx, createDocx, withNotePartsFromSource } =
    await import("./rezip");
  // Which note paragraphs changed is the session's journal, as for the body,
  // never a diff of the notes' re-serialization: a note the session never
  // touched (one the package held without a reference, say) is written as the
  // package held it, even where the model cannot spell its source markup.
  // Every paragraph of a note counts, at any depth, and one in a text box
  // drawn in a note is the note's too. Note parts are written against the
  // opened package: a note paragraph an edit and its undo left as the source
  // held it is no change. Against a later save it is one, so such a save
  // writes its note parts against the opened package once it is made.
  const noteChanges = snapshot.noteParagraphs;
  const sourceNoteChanges = new Set(noteChanges.changedSinceSource.map(idKey));
  const unchangedNotes = new Set(
    noteChanges.touched.map(idKey).filter((id) => !sourceNoteChanges.has(id)),
  );
  // So is a paragraph of any story that the session left as the source held
  // it, which a save against the opened package writes as the source's bytes.
  const restored = new Set(snapshot.restoredSinceSource.map(idKey));
  const changedBlockIds =
    snapshot.saveBaseline === undefined
      ? snapshot.changedBlockIds.filter(
          (id) => !unchangedNotes.has(idKey(id)) && !restored.has(idKey(id)),
        )
      : snapshot.changedBlockIds;
  const changedNoteParaIds = new Set(
    snapshot.saveBaseline === undefined
      ? noteChanges.changedSinceSource
      : noteChanges.touched,
  );
  const useSelective =
    flags.selectiveSave && options?.mode !== FOLIO_DOCX_SERIALIZATION_MODE.full;
  let selectiveBuffer: ArrayBuffer | null = null;
  if (baseline && (useSelective || flags.selectiveSaveTripwire)) {
    const { attemptSelectiveSave } = await import("./selectiveSave");
    selectiveBuffer = await attemptSelectiveSave(written, baseline, {
      bodyAuthority: "canonical",
      changedParaIds: new Set([
        ...changedBlockIds.filter((id) => !heldBack.has(idKey(id))),
        ...rejectionRewrote,
      ]),
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
  const path = fullBuffer === buffer ? "full-repack" : "selective";
  const source = document.originalBuffer;
  if (
    snapshot.saveBaseline !== undefined &&
    source !== undefined &&
    noteChanges.touched.length > 0
  )
    buffer = await withNotePartsFromSource({
      document: written,
      source,
      saved: buffer,
      changedNoteParaIds: new Set(noteChanges.changedSinceSource),
    });
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
  options?.onSerializationPath?.(path);
  return {
    // A part the session left as it read it is written as the package held
    // it, without the ids normalizing the package gave it.
    buffer: await withNormalizedSourceParts(buffer, snapshot.document),
    path,
    document,
    version: snapshot.version,
    diagnostics,
    tripwireResult,
  };
};
