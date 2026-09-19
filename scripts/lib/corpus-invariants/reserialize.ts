/**
 * Make the real serializers run.
 *
 * folio does not re-serialize what it has not changed. A paragraph's `w:pPr`,
 * a table's `w:tblPr`, a header's whole part, a drawing and a content control's
 * properties are each replayed from bytes captured at parse time whenever a
 * fingerprint says the model still agrees with them. That is correct for
 * fidelity and it is the reason the existing invariants prove less than they
 * appear to: a round trip over an untouched document exercises the capture
 * machinery, not the serializers, so a serializer that writes `left` where the
 * source said `start`, or drops a property the model holds, passes every one of
 * them. Every edited document takes the other path, so those bugs are live.
 *
 * This invariant removes the captures and compares what the serializers then
 * produce against what replay produced. A difference is a serializer defect
 * that verbatim replay was hiding.
 *
 * Only slots with a model fallback are stripped. A drawing in `preserveOnly`
 * mode and a shape's fill or outline markup have no model behind them: their
 * captured XML is the content, and removing it would test deletion rather than
 * serialization.
 */

import { parseDocx } from "@stll/folio-core/docx/parser";
import { repackDocx } from "@stll/folio-core/docx/rezip";
import type { Document } from "@stll/folio-core/types/document";
import { Result } from "better-result";

import { failureFromError } from "../corpus-signature";
import {
  type CorpusInvariantInput,
  type CorpusInvariantOutcome,
  EXTENDED_CORPUS_INVARIANTS,
  timeStage,
} from "./contract";
import {
  describePackageDifferences,
  differenceFailures,
  type PackageDifferences,
} from "./model-equality";

/**
 * What to do with each capture slot the model can rebuild.
 *
 * `strip` removes the slot so the serializer must build the element from the
 * model. `poison` replaces a fingerprint with a value that can never match,
 * which is the same path an edited drawing takes. `keep` marks a slot that is
 * itself the content.
 */
const CAPTURE_POLICIES = {
  sourceXml: "strip",
  gridSourceXml: "strip",
  verbatimXml: "strip",
  verbatimFingerprint: "strip",
  rawPropertiesXml: "strip",
  rawEndPropertiesXml: "strip",
  rawImageFingerprint: "poison",
  /** A `preserveOnly` drawing, a shape fill and a shape outline have no model behind them. */
  rawXml: "keep",
  rawXmlMode: "keep",
  rawWatermarkXml: "keep",
  /** Block-level range markers: the captured markup is the only model there is. */
  rawMarkersBefore: "keep",
  rawMarkersAfter: "keep",
} as const;

type CaptureSlot = keyof typeof CAPTURE_POLICIES;

/** `EDITED_PREVIEW_FINGERPRINT` in `packages/core/src/docx/imageRawXml.ts`. */
const EDITED_PREVIEW_FINGERPRINT = "editedPreview";

const CAPTURE_SLOTS = new Set<string>(Object.keys(CAPTURE_POLICIES));

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

/**
 * Remove every capture slot in place, depth first.
 *
 * The caller owns a private clone, so mutation is safe and a copying walk would
 * only double the peak memory of a package that is already large.
 */
const stripCapturesInPlace = (value: unknown, seen: WeakSet<object>): void => {
  if (Array.isArray(value)) {
    for (const item of value) {
      stripCapturesInPlace(item, seen);
    }
    return;
  }
  if (value instanceof Map) {
    for (const entry of value.values()) {
      stripCapturesInPlace(entry, seen);
    }
    return;
  }
  if (!isRecord(value) || seen.has(value)) {
    return;
  }
  seen.add(value);
  for (const key of Object.keys(value)) {
    if (!CAPTURE_SLOTS.has(key)) {
      stripCapturesInPlace(value[key], seen);
      continue;
    }
    // SAFETY: the key was just proven a member of the policy map's key set.
    const policy = CAPTURE_POLICIES[key as CaptureSlot];
    if (policy === "strip") {
      // Every replay gate reads the slot and tests it against `undefined`, so
      // clearing it defeats replay exactly as removing the key would, without
      // deleting a computed property.
      value[key] = undefined;
      continue;
    }
    if (policy === "poison") {
      value[key] = EDITED_PREVIEW_FINGERPRINT;
    }
  }
};

/**
 * A document whose every rebuildable capture is gone.
 *
 * `structuredClone` does the load-bearing half: paragraph property captures
 * live in a `WeakMap` keyed by the paragraph object, and the repository keeps
 * the only clone that carries them across (`@stll/folio-core/docx/document-clone`)
 * deliberately separate. A plain structural clone therefore arrives without
 * them, and the walk removes the slots the model itself holds.
 */
export const withoutSerializerCaptures = (document: Document): Document => {
  const cloned = structuredClone(document);
  stripCapturesInPlace(cloned.package, new WeakSet());
  return cloned;
};

const REPLAY_PREFIX = "replay hides a serializer difference at";
const REPACK_PREFIX = "a full repack already loses";

const parseBuffer = (buffer: ArrayBuffer): Promise<Document> =>
  parseDocx(buffer, { preloadFonts: false });

const save = (document: Document): Promise<ArrayBuffer> =>
  repackDocx(document, { updateModifiedDate: false });

/**
 * The differences a plain repack shows too, taken only when the forced run
 * found any.
 *
 * A difference the control repack also produces is not about replay, and
 * charging every passing file a second repack to learn that would double the
 * invariant's cost for nothing. The answer is per difference rather than per
 * file: one file can lose a run's formatting to the repack and a drawing's
 * outline only to the forced path, and prefixing both the same way would send
 * a reader to the wrong module for one of them.
 */
const differencesSurvivingPlainRepack = async (parsed: Document): Promise<ReadonlySet<string>> => {
  const control = await Result.tryPromise({
    try: async () => parseBuffer(await save(parsed)),
    catch: (cause: unknown) => cause,
  });
  if (control.isErr()) {
    return new Set();
  }
  return new Set(describePackageDifferences(parsed, control.value).messages);
};

export const runReserializeInvariant = async ({
  parsed,
}: CorpusInvariantInput): Promise<CorpusInvariantOutcome> => {
  const timings: Record<string, number> = {};

  const stripped = await timeStage(timings, "strip-captures", () =>
    Result.try(() => withoutSerializerCaptures(parsed)),
  );
  if (stripped.isErr()) {
    return {
      failures: [failureFromError(EXTENDED_CORPUS_INVARIANTS.reserialize, stripped.error)],
      timings,
    };
  }

  const forced = await timeStage(timings, "forced-save", () =>
    Result.tryPromise({ try: () => save(stripped.value), catch: (cause: unknown) => cause }),
  );
  if (forced.isErr()) {
    return {
      failures: [failureFromError(EXTENDED_CORPUS_INVARIANTS.reserialize, forced.error)],
      timings,
    };
  }

  const reparsed = await timeStage(timings, "forced-parse", () =>
    Result.tryPromise({ try: () => parseBuffer(forced.value), catch: (cause: unknown) => cause }),
  );
  if (reparsed.isErr()) {
    return {
      failures: [failureFromError(EXTENDED_CORPUS_INVARIANTS.reserialize, reparsed.error)],
      timings,
    };
  }

  const differences: PackageDifferences = await timeStage(timings, "compare", () =>
    describePackageDifferences(parsed, reparsed.value),
  );
  if (differences.messages.length === 0 && differences.omitted === 0) {
    return { failures: [], timings };
  }

  const alsoWithoutStripping = await timeStage(timings, "control-save", () =>
    differencesSurvivingPlainRepack(parsed),
  );
  return {
    failures: differenceFailures(
      EXTENDED_CORPUS_INVARIANTS.reserialize,
      differences,
      (message) =>
        `${alsoWithoutStripping.has(message) ? REPACK_PREFIX : REPLAY_PREFIX} ${message}`,
    ),
    timings,
  };
};
