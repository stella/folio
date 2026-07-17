import { createHash } from "node:crypto";

import { panic, TaggedError } from "better-result";

import { getDocxXmlSafetyIssue } from "../xmlSafety";
import type { DocxArchiveOptions } from "./boundedArchive";
import { DocxArchiveError } from "./boundedArchive";
import {
  FolioDocxPackageInspectionError,
  inspectDocxPackage,
  type FolioDocxInspectedXmlPart,
} from "./inspectDocxPackage";

export const FOLIO_DOCX_XML_PATCH_PROPOSAL_VERSION = 1 as const;
export const FOLIO_DOCX_XML_PATCH_PROPOSAL_PROFILE = "folio-xml-patch-proposal-v1" as const;
export const FOLIO_DOCX_XML_PATCH_PROPOSAL_DEFAULTS = Object.freeze({
  maxReplacements: 16,
  maxPartBytes: 8 * 1024 * 1024,
  maxTotalBytes: 16 * 1024 * 1024,
} as const);

export const FOLIO_DOCX_XML_PATCH_PROPOSAL_ISSUE_CODES = Object.freeze([
  "invalid-proposal",
  "unsupported-version",
  "too-many-replacements",
  "duplicate-part",
  "invalid-part-path",
  "invalid-base-sha256",
  "part-not-allowed",
  "part-not-found",
  "part-not-xml",
  "replacement-too-large",
  "replacements-too-large",
  "xml-doctype-forbidden",
  "xml-not-well-formed",
  "xml-encoding-mismatch",
  "base-hash-mismatch",
  "package-inspection-failed",
] as const);

export type FolioDocxXmlPatchProposalIssueCode =
  (typeof FOLIO_DOCX_XML_PATCH_PROPOSAL_ISSUE_CODES)[number];

export type FolioDocxXmlReplacement = {
  readonly path: string;
  readonly baseSha256: string;
  readonly replacementXml: string;
};

export type FolioDocxXmlPatchProposal = {
  readonly version: typeof FOLIO_DOCX_XML_PATCH_PROPOSAL_VERSION;
  readonly replacements: readonly [FolioDocxXmlReplacement, ...FolioDocxXmlReplacement[]];
};

export type FolioDocxXmlPatchProposalLimits = {
  readonly maxReplacements?: number;
  readonly maxPartBytes?: number;
  readonly maxTotalBytes?: number;
};

export type EvaluateDocxXmlPatchProposalArgs = {
  readonly bytes: ArrayBuffer | Uint8Array;
  readonly proposal: unknown;
  /** Exact package paths authorized by the server-side caller. */
  readonly allowedParts: readonly string[];
  readonly archive?: DocxArchiveOptions;
  readonly limits?: FolioDocxXmlPatchProposalLimits;
};

export type FolioDocxXmlPatchProposalIssue = {
  readonly code: FolioDocxXmlPatchProposalIssueCode;
  readonly message: string;
  readonly proposalPath?: string;
  readonly part?: string;
};

export type FolioDocxPreparedXmlReplacement = {
  readonly path: string;
  readonly baseSha256: string;
  readonly currentSha256: string;
  readonly replacementSha256: string;
  readonly replacementByteLength: number;
  readonly encoding: "utf-8";
};

export type FolioDocxXmlPatchProposalEvaluation =
  | {
      readonly version: typeof FOLIO_DOCX_XML_PATCH_PROPOSAL_VERSION;
      readonly profile: typeof FOLIO_DOCX_XML_PATCH_PROPOSAL_PROFILE;
      readonly status: "accepted";
      readonly producesOutput: false;
      readonly issues: readonly [];
      readonly replacements: readonly [
        FolioDocxPreparedXmlReplacement,
        ...FolioDocxPreparedXmlReplacement[],
      ];
      readonly limits: Required<FolioDocxXmlPatchProposalLimits>;
    }
  | {
      readonly version: typeof FOLIO_DOCX_XML_PATCH_PROPOSAL_VERSION;
      readonly profile: typeof FOLIO_DOCX_XML_PATCH_PROPOSAL_PROFILE;
      readonly status: "rejected";
      readonly producesOutput: false;
      readonly issues: readonly [
        FolioDocxXmlPatchProposalIssue,
        ...FolioDocxXmlPatchProposalIssue[],
      ];
      readonly replacements: readonly [];
      readonly limits: Required<FolioDocxXmlPatchProposalLimits>;
    };

export class InvalidFolioDocxXmlPatchProposalError extends TaggedError(
  "InvalidFolioDocxXmlPatchProposalError",
)<{
  message: string;
  path: string;
  reason: string;
}>() {}

export class InvalidFolioDocxXmlPatchProposalOptionsError extends TaggedError(
  "InvalidFolioDocxXmlPatchProposalOptionsError",
)<{
  message: string;
  path: string;
}>() {}

const SHA256_PATTERN = /^[\da-f]{64}$/u;
const XML_DECLARED_ENCODING_PATTERN =
  /^\uFEFF?<\?xml[^>]*\bencoding\s*=\s*["']([^"']+)["'][^>]*\?>/iu;
const MAX_PACKAGE_PATH_CODE_UNITS = 1024;

type PrepareXmlReplacementOptions = {
  replacement: FolioDocxXmlReplacement;
  encodedByPath: ReadonlyMap<string, Uint8Array>;
  inspectedByPath: ReadonlyMap<string, FolioDocxInspectedXmlPart>;
};

const prepareXmlReplacement = ({
  replacement,
  encodedByPath,
  inspectedByPath,
}: PrepareXmlReplacementOptions): FolioDocxPreparedXmlReplacement => {
  const replacementBytes = encodedByPath.get(replacement.path);
  const current = inspectedByPath.get(replacement.path);
  if (replacementBytes === undefined || current === undefined) {
    return panic("Accepted XML replacement is missing prepared package data");
  }
  return {
    path: replacement.path,
    baseSha256: replacement.baseSha256,
    currentSha256: current.sha256,
    replacementSha256: createHash("sha256").update(replacementBytes).digest("hex"),
    replacementByteLength: replacementBytes.byteLength,
    encoding: "utf-8",
  };
};

const isPlainObject = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

const invalidProposal = (path: string, reason: string): never => {
  throw new InvalidFolioDocxXmlPatchProposalError({
    message: `Invalid XML patch proposal at ${path}: ${reason}.`,
    path,
    reason,
  });
};

const assertAllowedKeys = (
  value: Record<string, unknown>,
  path: string,
  allowedKeys: readonly string[],
): void => {
  const allowed = new Set(allowedKeys);
  const unexpected = Object.keys(value).find((key) => !allowed.has(key));
  if (unexpected !== undefined) {
    invalidProposal(`${path}.${unexpected}`, "unexpected property");
  }
};

const readString = (value: Record<string, unknown>, key: string, path: string): string => {
  const result = value[key];
  if (typeof result !== "string") {
    return invalidProposal(`${path}.${key}`, "expected a string");
  }
  return result;
};

const parseXmlReplacement = (value: unknown, index: number): FolioDocxXmlReplacement => {
  const path = `$.replacements[${index}]`;
  if (!isPlainObject(value)) {
    return invalidProposal(path, "expected an object");
  }
  assertAllowedKeys(value, path, ["path", "baseSha256", "replacementXml"]);
  return {
    path: readString(value, "path", path),
    baseSha256: readString(value, "baseSha256", path),
    replacementXml: readString(value, "replacementXml", path),
  };
};

export const parseFolioDocxXmlPatchProposal = (value: unknown): FolioDocxXmlPatchProposal => {
  if (!isPlainObject(value)) {
    return invalidProposal("$", "expected an object");
  }
  assertAllowedKeys(value, "$", ["version", "replacements"]);
  const version = value["version"];
  if (typeof version !== "number" || !Number.isSafeInteger(version)) {
    return invalidProposal("$.version", "expected an integer contract version");
  }
  if (version !== FOLIO_DOCX_XML_PATCH_PROPOSAL_VERSION) {
    return invalidProposal("$.version", "unsupported contract version");
  }
  const replacements = value["replacements"];
  if (!Array.isArray(replacements) || replacements.length === 0) {
    return invalidProposal("$.replacements", "expected a non-empty array");
  }
  if (replacements.length > FOLIO_DOCX_XML_PATCH_PROPOSAL_DEFAULTS.maxReplacements) {
    return invalidProposal(
      "$.replacements",
      `expected at most ${FOLIO_DOCX_XML_PATCH_PROPOSAL_DEFAULTS.maxReplacements} replacements`,
    );
  }

  const firstReplacement = replacements.at(0);
  if (firstReplacement === undefined) {
    return invalidProposal("$.replacements", "expected a non-empty array");
  }
  return {
    version: FOLIO_DOCX_XML_PATCH_PROPOSAL_VERSION,
    replacements: [
      parseXmlReplacement(firstReplacement, 0),
      ...replacements
        .slice(1)
        .map((replacement, index) => parseXmlReplacement(replacement, index + 1)),
    ],
  };
};

type ResolveLimitOptions = {
  value: number | undefined;
  fallback: number;
  path: string;
};

const resolveLimit = ({ value, fallback, path }: ResolveLimitOptions): number => {
  const limit = value ?? fallback;
  if (!Number.isSafeInteger(limit) || limit < 0) {
    throw new InvalidFolioDocxXmlPatchProposalOptionsError({
      message: `${path} must be a non-negative safe integer.`,
      path,
    });
  }
  return limit;
};

const resolveLimits = (
  limits: FolioDocxXmlPatchProposalLimits | undefined,
): Required<FolioDocxXmlPatchProposalLimits> => {
  const maxReplacements = resolveLimit({
    value: limits?.maxReplacements,
    fallback: FOLIO_DOCX_XML_PATCH_PROPOSAL_DEFAULTS.maxReplacements,
    path: "limits.maxReplacements",
  });
  if (maxReplacements > FOLIO_DOCX_XML_PATCH_PROPOSAL_DEFAULTS.maxReplacements) {
    throw new InvalidFolioDocxXmlPatchProposalOptionsError({
      message: `limits.maxReplacements must not exceed ${FOLIO_DOCX_XML_PATCH_PROPOSAL_DEFAULTS.maxReplacements}.`,
      path: "limits.maxReplacements",
    });
  }
  return {
    maxReplacements,
    maxPartBytes: resolveLimit({
      value: limits?.maxPartBytes,
      fallback: FOLIO_DOCX_XML_PATCH_PROPOSAL_DEFAULTS.maxPartBytes,
      path: "limits.maxPartBytes",
    }),
    maxTotalBytes: resolveLimit({
      value: limits?.maxTotalBytes,
      fallback: FOLIO_DOCX_XML_PATCH_PROPOSAL_DEFAULTS.maxTotalBytes,
      path: "limits.maxTotalBytes",
    }),
  };
};

const isNormalizedPackagePath = (path: string): boolean =>
  path.length > 0 &&
  path.length <= MAX_PACKAGE_PATH_CODE_UNITS &&
  !path.startsWith("/") &&
  !path.endsWith("/") &&
  !path.includes("\\") &&
  path.split("/").every((segment) => segment.length > 0 && segment !== "." && segment !== "..");

const resolveAllowedParts = (allowedParts: readonly unknown[]): ReadonlySet<string> => {
  const resolved = new Set<string>();
  for (const [index, path] of allowedParts.entries()) {
    if (typeof path !== "string" || !isNormalizedPackagePath(path)) {
      throw new InvalidFolioDocxXmlPatchProposalOptionsError({
        message: `allowedParts[${index}] must be a normalized package path.`,
        path: `allowedParts[${index}]`,
      });
    }
    if (resolved.has(path)) {
      throw new InvalidFolioDocxXmlPatchProposalOptionsError({
        message: `allowedParts[${index}] must be unique.`,
        path: `allowedParts[${index}]`,
      });
    }
    resolved.add(path);
  }
  return resolved;
};

type RejectedEvaluationOptions = {
  issues: readonly [FolioDocxXmlPatchProposalIssue, ...FolioDocxXmlPatchProposalIssue[]];
  limits: Required<FolioDocxXmlPatchProposalLimits>;
};

const rejectedEvaluation = ({
  issues,
  limits,
}: RejectedEvaluationOptions): FolioDocxXmlPatchProposalEvaluation => ({
  version: FOLIO_DOCX_XML_PATCH_PROPOSAL_VERSION,
  profile: FOLIO_DOCX_XML_PATCH_PROPOSAL_PROFILE,
  status: "rejected",
  producesOutput: false,
  issues,
  replacements: [],
  limits,
});

type AcceptedEvaluationOptions = {
  replacements: readonly [FolioDocxPreparedXmlReplacement, ...FolioDocxPreparedXmlReplacement[]];
  limits: Required<FolioDocxXmlPatchProposalLimits>;
};

const acceptedEvaluation = ({
  replacements,
  limits,
}: AcceptedEvaluationOptions): FolioDocxXmlPatchProposalEvaluation => ({
  version: FOLIO_DOCX_XML_PATCH_PROPOSAL_VERSION,
  profile: FOLIO_DOCX_XML_PATCH_PROPOSAL_PROFILE,
  status: "accepted",
  producesOutput: false,
  issues: [],
  replacements,
  limits,
});

const inspectionIssue = (
  error: FolioDocxPackageInspectionError,
): FolioDocxXmlPatchProposalIssue => {
  if (error.code === "part-not-found" || error.code === "part-not-xml") {
    return {
      code: error.code,
      message: error.message,
      ...(error.part !== undefined && { part: error.part }),
    };
  }
  return {
    code: "package-inspection-failed",
    message: "The package could not be inspected within the configured limits.",
    ...(error.part !== undefined && { part: error.part }),
  };
};

export const evaluateDocxXmlPatchProposal = async ({
  bytes,
  proposal: rawProposal,
  allowedParts,
  archive,
  limits: rawLimits,
}: EvaluateDocxXmlPatchProposalArgs): Promise<FolioDocxXmlPatchProposalEvaluation> => {
  const limits = resolveLimits(rawLimits);
  const allowed = resolveAllowedParts(allowedParts);
  const rawReplacements = isPlainObject(rawProposal) ? rawProposal["replacements"] : undefined;
  if (
    isPlainObject(rawProposal) &&
    rawProposal["version"] === FOLIO_DOCX_XML_PATCH_PROPOSAL_VERSION &&
    Array.isArray(rawReplacements) &&
    rawReplacements.length > limits.maxReplacements
  ) {
    return rejectedEvaluation({
      issues: [
        {
          code: "too-many-replacements",
          message: `Proposal contains ${rawReplacements.length} replacements (max ${limits.maxReplacements}).`,
          proposalPath: "$.replacements",
        },
      ],
      limits,
    });
  }
  let proposal: FolioDocxXmlPatchProposal;
  try {
    proposal = parseFolioDocxXmlPatchProposal(rawProposal);
  } catch (error) {
    if (!(error instanceof InvalidFolioDocxXmlPatchProposalError)) {
      throw error;
    }
    return rejectedEvaluation({
      issues: [
        {
          code:
            error.path === "$.version" &&
            isPlainObject(rawProposal) &&
            typeof rawProposal["version"] === "number" &&
            Number.isSafeInteger(rawProposal["version"])
              ? "unsupported-version"
              : "invalid-proposal",
          message: error.message,
          proposalPath: error.path,
        },
      ],
      limits,
    });
  }

  const issues: FolioDocxXmlPatchProposalIssue[] = [];
  const seen = new Set<string>();
  const encodedByPath = new Map<string, Uint8Array>();
  let encodedTotalBytes = 0;
  let totalCodeUnits = 0;
  let totalLimitReported = false;
  for (const [index, replacement] of proposal.replacements.entries()) {
    const proposalPath = `$.replacements[${index}]`;
    if (seen.has(replacement.path)) {
      issues.push({
        code: "duplicate-part",
        message: "Each package part may appear only once in a proposal.",
        proposalPath: `${proposalPath}.path`,
        part: replacement.path,
      });
    }
    seen.add(replacement.path);
    if (!isNormalizedPackagePath(replacement.path)) {
      issues.push({
        code: "invalid-part-path",
        message: "Replacement paths must be normalized package paths.",
        proposalPath: `${proposalPath}.path`,
        part: replacement.path,
      });
    } else if (!allowed.has(replacement.path)) {
      issues.push({
        code: "part-not-allowed",
        message: "The server-side policy does not allow replacing this package part.",
        proposalPath: `${proposalPath}.path`,
        part: replacement.path,
      });
    }
    if (replacement.baseSha256.length !== 64 || !SHA256_PATTERN.test(replacement.baseSha256)) {
      issues.push({
        code: "invalid-base-sha256",
        message: "baseSha256 must be a lowercase hexadecimal SHA-256 digest.",
        proposalPath: `${proposalPath}.baseSha256`,
        part: replacement.path,
      });
    }

    const replacementCodeUnits = replacement.replacementXml.length;
    const exceedsPartLowerBound = replacementCodeUnits > limits.maxPartBytes;
    if (exceedsPartLowerBound) {
      issues.push({
        code: "replacement-too-large",
        message: `Replacement XML exceeds the ${limits.maxPartBytes}-byte per-part limit.`,
        proposalPath: `${proposalPath}.replacementXml`,
        part: replacement.path,
      });
    }
    totalCodeUnits += replacementCodeUnits;
    if (totalCodeUnits > limits.maxTotalBytes && !totalLimitReported) {
      totalLimitReported = true;
      issues.push({
        code: "replacements-too-large",
        message: `Replacement XML exceeds the ${limits.maxTotalBytes}-byte cumulative limit.`,
        proposalPath: "$.replacements",
      });
    }
    if (exceedsPartLowerBound || totalLimitReported) {
      continue;
    }

    const replacementBytes = new TextEncoder().encode(replacement.replacementXml);
    const exceedsEncodedPart = replacementBytes.byteLength > limits.maxPartBytes;
    const nextEncodedTotalBytes = encodedTotalBytes + replacementBytes.byteLength;
    if (exceedsEncodedPart) {
      issues.push({
        code: "replacement-too-large",
        message: `Replacement XML exceeds the ${limits.maxPartBytes}-byte per-part limit.`,
        proposalPath: `${proposalPath}.replacementXml`,
        part: replacement.path,
      });
    }
    if (nextEncodedTotalBytes > limits.maxTotalBytes) {
      totalLimitReported = true;
      issues.push({
        code: "replacements-too-large",
        message: `Replacement XML exceeds the ${limits.maxTotalBytes}-byte cumulative limit.`,
        proposalPath: "$.replacements",
      });
    }
    encodedTotalBytes = nextEncodedTotalBytes;
    if (exceedsEncodedPart || totalLimitReported) {
      continue;
    }
    encodedByPath.set(replacement.path, replacementBytes);

    const safetyIssue = getDocxXmlSafetyIssue(replacement.replacementXml);
    if (safetyIssue === "doctype-forbidden") {
      issues.push({
        code: "xml-doctype-forbidden",
        message: "Replacement XML must not declare a document type.",
        proposalPath: `${proposalPath}.replacementXml`,
        part: replacement.path,
      });
    } else if (safetyIssue === "not-well-formed") {
      issues.push({
        code: "xml-not-well-formed",
        message: "Replacement XML must be well formed.",
        proposalPath: `${proposalPath}.replacementXml`,
        part: replacement.path,
      });
    } else {
      const declaredEncoding = replacement.replacementXml
        .match(XML_DECLARED_ENCODING_PATTERN)
        ?.at(1);
      if (
        declaredEncoding !== undefined &&
        declaredEncoding.toLowerCase() !== "utf-8" &&
        declaredEncoding.toLowerCase() !== "utf8"
      ) {
        issues.push({
          code: "xml-encoding-mismatch",
          message: "Replacement XML declarations must use UTF-8 encoding.",
          proposalPath: `${proposalPath}.replacementXml`,
          part: replacement.path,
        });
      }
    }
  }

  if (issues.length > 0) {
    const firstIssue = issues.at(0);
    if (firstIssue === undefined) {
      return panic("Rejected XML patch proposal is missing an issue");
    }
    return rejectedEvaluation({ issues: [firstIssue, ...issues.slice(1)], limits });
  }

  let inspected;
  try {
    inspected = await inspectDocxPackage(bytes, {
      ...(archive === undefined ? {} : { archive }),
      xmlParts: proposal.replacements.map(({ path }) => path),
      limits: {
        maxXmlParts: limits.maxReplacements,
        maxXmlPartBytes: limits.maxPartBytes,
        maxXmlTotalBytes: limits.maxTotalBytes,
      },
    });
  } catch (error) {
    if (error instanceof FolioDocxPackageInspectionError) {
      return rejectedEvaluation({
        issues: [inspectionIssue(error)],
        limits,
      });
    }
    if (error instanceof DocxArchiveError) {
      return rejectedEvaluation({
        issues: [
          {
            code: "package-inspection-failed",
            message: "The package could not be inspected within the configured limits.",
          },
        ],
        limits,
      });
    }
    throw error;
  }

  const inspectedByPath = new Map(inspected.xmlParts.map((part) => [part.path, part] as const));
  for (const [index, replacement] of proposal.replacements.entries()) {
    const current = inspectedByPath.get(replacement.path);
    if (current === undefined || current.sha256 === replacement.baseSha256) {
      continue;
    }
    issues.push({
      code: "base-hash-mismatch",
      message: "The package part changed after the proposal base was inspected.",
      proposalPath: `$.replacements[${index}].baseSha256`,
      part: replacement.path,
    });
  }
  if (issues.length > 0) {
    const firstIssue = issues.at(0);
    if (firstIssue === undefined) {
      return panic("Rejected XML patch proposal is missing an issue");
    }
    return rejectedEvaluation({ issues: [firstIssue, ...issues.slice(1)], limits });
  }

  const preparedReplacements = proposal.replacements.map((replacement) =>
    prepareXmlReplacement({ replacement, encodedByPath, inspectedByPath }),
  );
  const firstPreparedReplacement = preparedReplacements.at(0);
  if (firstPreparedReplacement === undefined) {
    return panic("Accepted XML patch proposal is missing a prepared replacement");
  }
  return acceptedEvaluation({
    replacements: [firstPreparedReplacement, ...preparedReplacements.slice(1)],
    limits,
  });
};
