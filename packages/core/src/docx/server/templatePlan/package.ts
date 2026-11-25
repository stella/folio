/** A DOCX package opened for reading: its parts, and the parts its main document names. */
import { Result } from "better-result";
import { createHash } from "node:crypto";

import type { RelationshipMap } from "../../../types/document";
import {
  parseRelationships,
  RELATIONSHIP_TYPES,
  resolveRelativePath,
} from "../../relsParser";
import { loadDocxArchive } from "../boundedArchive";
import { TemplatePlanError } from "./types";

export const sha256Hex = (bytes: Uint8Array | string): string =>
  createHash("sha256").update(bytes).digest("hex");

export const relationshipsPathFor = (partPath: string): string => {
  const slash = partPath.lastIndexOf("/");
  const directory = slash === -1 ? "" : `${partPath.slice(0, slash)}/`;
  return `${directory}_rels/${partPath.slice(slash + 1)}.rels`;
};

export type DocumentPartKind =
  "styles" | "numbering" | "footnotes" | "endnotes" | "settings" | "comments";

export type LoadedPackage = {
  bytes: Uint8Array;
  sha256: string;
  paths: ReadonlySet<string>;
  readText: (path: string) => Promise<string | null>;
  readBytes: (path: string) => Promise<Uint8Array | null>;
  documentPath: string;
  documentXml: string;
  documentRelationshipsPath: string;
  documentRelationships: RelationshipMap;
  /** The path of the part the main document names for `kind`, when it names one. */
  partPath: (kind: DocumentPartKind) => string | null;
  /** The package path a relationship of the main document points at. */
  resolveDocumentTarget: (target: string) => string;
};

const PACKAGE_RELATIONSHIPS_PATH = "_rels/.rels";
const DEFAULT_DOCUMENT_PATH = "word/document.xml";

/** Open `bytes` within folio's archive and XML limits. */
export const loadPackage = async (
  bytes: Uint8Array,
): Promise<Result<LoadedPackage, TemplatePlanError>> => {
  const opened = await Result.tryPromise({
    try: async () => {
      const archive = await loadDocxArchive(bytes);
      const paths = new Set(archive.entries);
      const packageRelationships = parseRelationships(
        (await archive.readEntryString(PACKAGE_RELATIONSHIPS_PATH)) ?? "",
      );
      const officeDocument = [...packageRelationships.values()].find(
        ({ type }) => type === RELATIONSHIP_TYPES.officeDocument,
      );
      const documentPath =
        officeDocument === undefined
          ? DEFAULT_DOCUMENT_PATH
          : resolveRelativePath(
              PACKAGE_RELATIONSHIPS_PATH,
              officeDocument.target,
            );
      const documentXml = await archive.readEntryString(documentPath);
      const documentRelationshipsPath = relationshipsPathFor(documentPath);
      const documentRelationships = parseRelationships(
        (await archive.readEntryString(documentRelationshipsPath)) ?? "",
      );
      return {
        archive,
        paths,
        documentPath,
        documentXml,
        documentRelationshipsPath,
        documentRelationships,
      };
    },
    catch: (cause) =>
      new TemplatePlanError({
        message: "The DOCX package could not be read.",
        code: "archive-unreadable",
        cause,
      }),
  });
  if (Result.isError(opened)) {
    return opened;
  }
  const {
    archive,
    paths,
    documentPath,
    documentXml,
    documentRelationshipsPath,
    documentRelationships,
  } = opened.value;
  if (documentXml === null) {
    return Result.err(
      new TemplatePlanError({
        message: "The DOCX package has no main document part.",
        code: "part-missing",
        part: documentPath,
      }),
    );
  }
  const resolveDocumentTarget = (target: string): string =>
    resolveRelativePath(documentRelationshipsPath, target);
  return Result.ok({
    bytes,
    sha256: sha256Hex(bytes),
    paths,
    readText: async (path) => await archive.readEntryString(path),
    readBytes: async (path) => await archive.readEntryUint8(path),
    documentPath,
    documentXml,
    documentRelationshipsPath,
    documentRelationships,
    partPath: (kind) => {
      const relationship = [...documentRelationships.values()].find(
        ({ type, targetMode }) =>
          type === RELATIONSHIP_TYPES[kind] && targetMode !== "External",
      );
      if (relationship === undefined) {
        return null;
      }
      const path = resolveDocumentTarget(relationship.target);
      return paths.has(path) ? path : null;
    },
    resolveDocumentTarget,
  });
};

/** Read an XML part, refusing one the archive could not hand out. */
export const readPartText = async (
  pkg: LoadedPackage,
  path: string | null,
): Promise<Result<string | null, TemplatePlanError>> => {
  if (path === null) {
    return Result.ok(null);
  }
  return await Result.tryPromise({
    try: async () => await pkg.readText(path),
    catch: (cause) =>
      new TemplatePlanError({
        message: `The part ${path} could not be read.`,
        code: "xml-unreadable",
        part: path,
        cause,
      }),
  });
};

/**
 * Run one public entry of the module. Whatever a hostile or malformed
 * package makes a shared helper throw comes back as a typed refusal, so no
 * caller has to guard the call.
 */
export const guarded = async <Value>(
  run: () => Promise<Result<Value, TemplatePlanError>>,
): Promise<Result<Value, TemplatePlanError>> =>
  Result.flatten(
    await Result.tryPromise({
      try: run,
      catch: (cause) =>
        cause instanceof TemplatePlanError
          ? cause
          : new TemplatePlanError({
              message: "The package could not be processed.",
              code: "xml-unreadable",
              cause,
            }),
    }),
  );
