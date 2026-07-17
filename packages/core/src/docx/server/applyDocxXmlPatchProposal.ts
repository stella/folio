import { createHash } from "node:crypto";

import { panic, TaggedError } from "better-result";
import JSZip from "jszip";

import type { DocxArchiveOptions } from "./boundedArchive";
import { loadDocxArchive } from "./boundedArchive";
import {
  evaluateDocxXmlPatchProposal,
  FOLIO_DOCX_XML_PATCH_PROPOSAL_PROFILE,
  InvalidFolioDocxXmlPatchProposalError,
  parseFolioDocxXmlPatchProposal,
  type FolioDocxPreparedXmlReplacement,
  type FolioDocxXmlPatchProposal,
  type FolioDocxXmlPatchProposalEvaluation,
  type FolioDocxXmlPatchProposalLimits,
} from "./evaluateDocxXmlPatchProposal";
import {
  FOLIO_DOCX_CONFORMANCE_PROFILE,
  validateDocxConformance,
  type FolioDocxConformanceReport,
} from "./validateDocxConformance";

export const FOLIO_DOCX_XML_PATCH_APPLICATION_VERSION = 1 as const;
export const FOLIO_DOCX_XML_PATCH_APPLICATION_PROFILE = "folio-xml-patch-application-v1" as const;

export type ApplyDocxXmlPatchProposalArgs = {
  readonly bytes: ArrayBuffer | Uint8Array;
  readonly proposal: unknown;
  /** Exact package paths authorized by the server-side caller. */
  readonly allowedParts: readonly string[];
  readonly validationProfile: typeof FOLIO_DOCX_CONFORMANCE_PROFILE;
  readonly archive?: DocxArchiveOptions;
  readonly limits?: FolioDocxXmlPatchProposalLimits;
};

export type FolioDocxXmlPatchApplicationReceipt = {
  readonly version: typeof FOLIO_DOCX_XML_PATCH_APPLICATION_VERSION;
  readonly profile: typeof FOLIO_DOCX_XML_PATCH_APPLICATION_PROFILE;
  readonly proposalProfile: typeof FOLIO_DOCX_XML_PATCH_PROPOSAL_PROFILE;
  readonly validationProfile: typeof FOLIO_DOCX_CONFORMANCE_PROFILE;
  readonly input: {
    readonly sha256: string;
    readonly byteLength: number;
  };
  readonly output: {
    readonly sha256: string;
    readonly byteLength: number;
  };
  readonly replacements: readonly [
    FolioDocxPreparedXmlReplacement,
    ...FolioDocxPreparedXmlReplacement[],
  ];
};

type AcceptedProposalEvaluation = Extract<
  FolioDocxXmlPatchProposalEvaluation,
  { status: "accepted" }
>;
type RejectedProposalEvaluation = Extract<
  FolioDocxXmlPatchProposalEvaluation,
  { status: "rejected" }
>;
type ConformantReport = FolioDocxConformanceReport & { readonly status: "conformant" };

export type FolioDocxXmlPatchApplication =
  | {
      readonly version: typeof FOLIO_DOCX_XML_PATCH_APPLICATION_VERSION;
      readonly profile: typeof FOLIO_DOCX_XML_PATCH_APPLICATION_PROFILE;
      readonly status: "proposal-rejected";
      readonly producesOutput: false;
      readonly evaluation: Extract<FolioDocxXmlPatchProposalEvaluation, { status: "rejected" }>;
      readonly conformance: null;
      readonly receipt: null;
    }
  | {
      readonly version: typeof FOLIO_DOCX_XML_PATCH_APPLICATION_VERSION;
      readonly profile: typeof FOLIO_DOCX_XML_PATCH_APPLICATION_PROFILE;
      readonly status: "output-rejected";
      readonly producesOutput: false;
      readonly evaluation: Extract<FolioDocxXmlPatchProposalEvaluation, { status: "accepted" }>;
      readonly conformance: FolioDocxConformanceReport;
      readonly receipt: null;
    }
  | {
      readonly version: typeof FOLIO_DOCX_XML_PATCH_APPLICATION_VERSION;
      readonly profile: typeof FOLIO_DOCX_XML_PATCH_APPLICATION_PROFILE;
      readonly status: "applied";
      readonly producesOutput: true;
      readonly evaluation: Extract<FolioDocxXmlPatchProposalEvaluation, { status: "accepted" }>;
      readonly conformance: FolioDocxConformanceReport & { readonly status: "conformant" };
      readonly receipt: FolioDocxXmlPatchApplicationReceipt;
      readonly bytes: Uint8Array;
    };

export class UnsupportedFolioDocxXmlPatchApplicationProfileError extends TaggedError(
  "UnsupportedFolioDocxXmlPatchApplicationProfileError",
)<{
  message: string;
  profile: unknown;
}>() {}

export class FolioDocxXmlPatchApplicationError extends TaggedError(
  "FolioDocxXmlPatchApplicationError",
)<{
  message: string;
  stage: "verify" | "load" | "replace" | "generate";
  part?: string;
  cause?: unknown;
}>() {}

const sha256 = (bytes: Uint8Array): string => createHash("sha256").update(bytes).digest("hex");

const snapshotInput = (bytes: ArrayBuffer | Uint8Array): Uint8Array =>
  Uint8Array.from(bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes));

const snapshotArchiveOptions = (
  archive: DocxArchiveOptions | undefined,
): DocxArchiveOptions | undefined => {
  if (archive === undefined) {
    return undefined;
  }
  return {
    ...(archive.maxInputBytes === undefined ? {} : { maxInputBytes: archive.maxInputBytes }),
    ...(archive.maxEntryBytes === undefined ? {} : { maxEntryBytes: archive.maxEntryBytes }),
    ...(archive.maxTotalBytes === undefined ? {} : { maxTotalBytes: archive.maxTotalBytes }),
    ...(archive.maxEntries === undefined ? {} : { maxEntries: archive.maxEntries }),
  };
};

const snapshotProposalLimits = (
  limits: FolioDocxXmlPatchProposalLimits | undefined,
): FolioDocxXmlPatchProposalLimits | undefined => {
  if (limits === undefined) {
    return undefined;
  }
  return {
    ...(limits.maxReplacements === undefined ? {} : { maxReplacements: limits.maxReplacements }),
    ...(limits.maxPartBytes === undefined ? {} : { maxPartBytes: limits.maxPartBytes }),
    ...(limits.maxTotalBytes === undefined ? {} : { maxTotalBytes: limits.maxTotalBytes }),
  };
};

type EvaluateProposalOptions = {
  bytes: Uint8Array;
  proposal: unknown;
  allowedParts: readonly string[];
  archive: DocxArchiveOptions | undefined;
  limits: FolioDocxXmlPatchProposalLimits | undefined;
};

const evaluateProposal = async ({
  bytes,
  proposal,
  allowedParts,
  archive,
  limits,
}: EvaluateProposalOptions): Promise<FolioDocxXmlPatchProposalEvaluation> =>
  await evaluateDocxXmlPatchProposal({
    bytes,
    proposal,
    allowedParts,
    ...(archive === undefined ? {} : { archive }),
    ...(limits === undefined ? {} : { limits }),
  });

type ProposalRejectedOptions = {
  evaluation: RejectedProposalEvaluation;
};

const proposalRejected = ({
  evaluation,
}: ProposalRejectedOptions): FolioDocxXmlPatchApplication => ({
  version: FOLIO_DOCX_XML_PATCH_APPLICATION_VERSION,
  profile: FOLIO_DOCX_XML_PATCH_APPLICATION_PROFILE,
  status: "proposal-rejected",
  producesOutput: false,
  evaluation,
  conformance: null,
  receipt: null,
});

const verifyArchiveContents = async (
  bytes: Uint8Array,
  options: DocxArchiveOptions | undefined,
): Promise<void> => {
  let archive;
  try {
    archive = await loadDocxArchive(bytes, options);
  } catch (cause) {
    throw new FolioDocxXmlPatchApplicationError({
      message: "The evaluated package could not be verified within the archive limits.",
      stage: "verify",
      cause,
    });
  }

  for (const entry of archive.entryMetadata.toSorted(({ path: left }, { path: right }) =>
    left < right ? -1 : Number(left > right),
  )) {
    if (entry.directory) {
      continue;
    }
    try {
      // oxlint-disable-next-line no-await-in-loop -- serialized reads enforce the archive budget
      const content = await archive.readEntryUint8(entry.path);
      if (content === null) {
        throw new FolioDocxXmlPatchApplicationError({
          message: "An archive entry disappeared during verification.",
          stage: "verify",
          part: entry.path,
        });
      }
    } catch (cause) {
      if (cause instanceof FolioDocxXmlPatchApplicationError) {
        throw cause;
      }
      throw new FolioDocxXmlPatchApplicationError({
        message: "An archive entry could not be verified within the configured limits.",
        stage: "verify",
        part: entry.path,
        cause,
      });
    }
  }
};

type ApplyReplacementsOptions = {
  bytes: Uint8Array;
  proposal: FolioDocxXmlPatchProposal;
};

const applyReplacements = async ({
  bytes,
  proposal,
}: ApplyReplacementsOptions): Promise<Uint8Array> => {
  let zip: JSZip;
  try {
    zip = await JSZip.loadAsync(bytes);
  } catch (cause) {
    throw new FolioDocxXmlPatchApplicationError({
      message: "The evaluated package could not be loaded for replacement.",
      stage: "load",
      cause,
    });
  }

  for (const replacement of proposal.replacements) {
    const current = zip.file(replacement.path);
    if (current === null) {
      throw new FolioDocxXmlPatchApplicationError({
        message: "An evaluated package part was unavailable during replacement.",
        stage: "replace",
      });
    }
    zip.file(replacement.path, new TextEncoder().encode(replacement.replacementXml), {
      binary: true,
      date: current.date,
      comment: current.comment,
      createFolders: false,
      unixPermissions: current.unixPermissions,
      dosPermissions: current.dosPermissions,
      compression: "DEFLATE",
      compressionOptions: { level: 6 },
    });
  }

  try {
    return await zip.generateAsync({
      type: "uint8array",
      compression: "DEFLATE",
      compressionOptions: { level: 6 },
    });
  } catch (cause) {
    throw new FolioDocxXmlPatchApplicationError({
      message: "The replacement package could not be generated.",
      stage: "generate",
      cause,
    });
  }
};

type OutputRejectedOptions = {
  evaluation: AcceptedProposalEvaluation;
  conformance: FolioDocxConformanceReport;
};

const outputRejected = ({
  evaluation,
  conformance,
}: OutputRejectedOptions): FolioDocxXmlPatchApplication => ({
  version: FOLIO_DOCX_XML_PATCH_APPLICATION_VERSION,
  profile: FOLIO_DOCX_XML_PATCH_APPLICATION_PROFILE,
  status: "output-rejected",
  producesOutput: false,
  evaluation,
  conformance,
  receipt: null,
});

type AppliedOptions = {
  input: Uint8Array;
  output: Uint8Array;
  evaluation: AcceptedProposalEvaluation;
  conformance: ConformantReport;
};

const applied = ({
  input,
  output,
  evaluation,
  conformance,
}: AppliedOptions): FolioDocxXmlPatchApplication => ({
  version: FOLIO_DOCX_XML_PATCH_APPLICATION_VERSION,
  profile: FOLIO_DOCX_XML_PATCH_APPLICATION_PROFILE,
  status: "applied",
  producesOutput: true,
  evaluation,
  conformance,
  receipt: {
    version: FOLIO_DOCX_XML_PATCH_APPLICATION_VERSION,
    profile: FOLIO_DOCX_XML_PATCH_APPLICATION_PROFILE,
    proposalProfile: FOLIO_DOCX_XML_PATCH_PROPOSAL_PROFILE,
    validationProfile: FOLIO_DOCX_CONFORMANCE_PROFILE,
    input: {
      sha256: sha256(input),
      byteLength: input.byteLength,
    },
    output: {
      sha256: sha256(output),
      byteLength: output.byteLength,
    },
    replacements: evaluation.replacements,
  },
  bytes: output,
});

/** Apply a guarded XML proposal and return output only after full profile validation. */
export const applyDocxXmlPatchProposal = async ({
  bytes,
  proposal: rawProposal,
  allowedParts,
  validationProfile,
  archive,
  limits,
}: ApplyDocxXmlPatchProposalArgs): Promise<FolioDocxXmlPatchApplication> => {
  if (validationProfile !== FOLIO_DOCX_CONFORMANCE_PROFILE) {
    throw new UnsupportedFolioDocxXmlPatchApplicationProfileError({
      message: "The requested package validation profile is not supported.",
      profile: validationProfile,
    });
  }

  const input = snapshotInput(bytes);
  const policyAllowedParts = [...allowedParts];
  const policyArchive = snapshotArchiveOptions(archive);
  const policyLimits = snapshotProposalLimits(limits);
  let proposal: FolioDocxXmlPatchProposal;
  try {
    proposal = parseFolioDocxXmlPatchProposal(rawProposal);
  } catch (error) {
    if (!(error instanceof InvalidFolioDocxXmlPatchProposalError)) {
      throw error;
    }
    const evaluation = await evaluateProposal({
      bytes: input,
      proposal: rawProposal,
      allowedParts: policyAllowedParts,
      archive: policyArchive,
      limits: policyLimits,
    });
    if (evaluation.status !== "rejected") {
      return panic("An invalid XML proposal produced an accepted evaluation");
    }
    return proposalRejected({ evaluation });
  }

  const evaluation = await evaluateProposal({
    bytes: input,
    proposal,
    allowedParts: policyAllowedParts,
    archive: policyArchive,
    limits: policyLimits,
  });
  if (evaluation.status === "rejected") {
    return proposalRejected({ evaluation });
  }

  await verifyArchiveContents(input, policyArchive);
  const output = await applyReplacements({ bytes: input, proposal });
  const conformance = await validateDocxConformance(output, {
    ...(policyArchive === undefined ? {} : { archive: policyArchive }),
  });
  if (conformance.status !== "conformant") {
    return outputRejected({ evaluation, conformance });
  }
  return applied({ input, output, evaluation, conformance });
};
