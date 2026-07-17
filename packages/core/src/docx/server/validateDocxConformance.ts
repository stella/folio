import { DOCX_CONFORMANCE_CLASSES } from "@stll/docx-core/model";
import { XMLValidator } from "fast-xml-parser";

import type { DocxConformanceClass } from "../../types/document";
import { detectDocxConformanceClass } from "../conformance";
import { DOCX_CONTAINER_TYPES, detectDocxContainerType } from "../encryption/containerFormat";
import { DocxModelValidationError, validateFolioDocumentModel } from "../modelValidation";
import { DocxParseError, parseDocx } from "../parser";
import {
  findChildByLocalName,
  getAttributeAnyPrefix,
  getLocalName,
  getNamespacePrefix,
  parseXmlDocument,
} from "../xmlParser";
import type { DocxArchive, DocxArchiveOptions } from "./boundedArchive";
import { DocxArchiveError, loadDocxArchive } from "./boundedArchive";

export const FOLIO_DOCX_CONFORMANCE_REPORT_VERSION = 1 as const;
export const FOLIO_DOCX_CONFORMANCE_PROFILE = "folio-supported-v1" as const;

export const FOLIO_DOCX_CONFORMANCE_CHECKS = Object.freeze([
  "archive-safety",
  "required-parts",
  "xml-well-formedness",
  "package-roots",
  "conformance-class",
  "canonical-model",
] as const);

export const FOLIO_DOCX_CONFORMANCE_ISSUE_CODES = Object.freeze([
  "archive-load-failed",
  "archive-input-too-large",
  "archive-too-many-entries",
  "archive-entry-too-large",
  "archive-total-too-large",
  "required-part-missing",
  "xml-doctype-forbidden",
  "xml-not-well-formed",
  "xml-read-failed",
  "required-xml-unreadable",
  "package-root-invalid",
  "conformance-class-unknown",
  "model-invalid",
  "model-warning",
  "parser-recovery",
  "parser-unsupported",
  "parser-failed",
  "encrypted-container",
  "container-not-zip",
] as const);

export type FolioDocxConformanceCheckId = (typeof FOLIO_DOCX_CONFORMANCE_CHECKS)[number];
export type FolioDocxConformanceIssueCode = (typeof FOLIO_DOCX_CONFORMANCE_ISSUE_CODES)[number];
export type FolioDocxConformanceCheckStatus = "passed" | "failed" | "indeterminate" | "not-run";
export type FolioDocxConformanceStatus = "invalid" | "conformant" | "indeterminate";

export type FolioDocxConformanceIssue = {
  readonly check: FolioDocxConformanceCheckId;
  readonly code: FolioDocxConformanceIssueCode;
  readonly message: string;
  readonly severity: "error" | "warning";
  readonly part?: string;
  readonly modelPath?: string;
  readonly count?: number;
};

export type FolioDocxConformanceCheck = {
  readonly id: FolioDocxConformanceCheckId;
  readonly status: FolioDocxConformanceCheckStatus;
};

export type FolioDocxConformanceReport = {
  readonly version: typeof FOLIO_DOCX_CONFORMANCE_REPORT_VERSION;
  readonly profile: typeof FOLIO_DOCX_CONFORMANCE_PROFILE;
  readonly status: FolioDocxConformanceStatus;
  readonly conformanceClass: DocxConformanceClass;
  readonly checks: readonly FolioDocxConformanceCheck[];
  readonly issues: readonly FolioDocxConformanceIssue[];
  readonly unverifiedStandardsDimensions: readonly [
    "complete-schema-constraints",
    "markup-compatibility-processing",
    "consumer-specific-rendering",
  ];
};

export type ValidateDocxConformanceOptions = {
  readonly archive?: DocxArchiveOptions;
};

const REQUIRED_PARTS = Object.freeze([
  "[Content_Types].xml",
  "_rels/.rels",
  "word/document.xml",
] as const);
const XML_PART_PATTERN = /(?:\.xml|\.rels)$/u;
const DOCUMENT_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
const CONTENT_TYPES_NAMESPACE = "http://schemas.openxmlformats.org/package/2006/content-types";
const RELATIONSHIPS_NAMESPACE = "http://schemas.openxmlformats.org/package/2006/relationships";
const DOCTYPE_PATTERN = /<!DOCTYPE(?:\s|>)/iu;
const UNVERIFIED_STANDARDS_DIMENSIONS = Object.freeze([
  "complete-schema-constraints",
  "markup-compatibility-processing",
  "consumer-specific-rendering",
] as const);

type MutableReport = {
  conformanceClass: DocxConformanceClass;
  checks: Map<FolioDocxConformanceCheckId, FolioDocxConformanceCheckStatus>;
  issues: FolioDocxConformanceIssue[];
};

const createMutableReport = (): MutableReport => ({
  conformanceClass: DOCX_CONFORMANCE_CLASSES.UNKNOWN,
  checks: new Map(FOLIO_DOCX_CONFORMANCE_CHECKS.map((id) => [id, "not-run"] as const)),
  issues: [],
});

const addIssue = (report: MutableReport, issue: FolioDocxConformanceIssue): void => {
  report.issues.push(issue);
};

const finishReport = (report: MutableReport): FolioDocxConformanceReport => {
  const checks = FOLIO_DOCX_CONFORMANCE_CHECKS.map((id) => ({
    id,
    status: report.checks.get(id) ?? "not-run",
  }));
  const hasFailedCheck = checks.some(({ status }) => status === "failed");
  const hasIndeterminateCheck = checks.some(({ status }) => status === "indeterminate");
  let status: FolioDocxConformanceStatus = "conformant";
  if (hasFailedCheck) {
    status = "invalid";
  } else if (
    hasIndeterminateCheck ||
    checks.some(({ status: checkStatus }) => checkStatus === "not-run")
  ) {
    status = "indeterminate";
  }

  return {
    version: FOLIO_DOCX_CONFORMANCE_REPORT_VERSION,
    profile: FOLIO_DOCX_CONFORMANCE_PROFILE,
    status,
    conformanceClass: report.conformanceClass,
    checks,
    issues: report.issues,
    unverifiedStandardsDimensions: UNVERIFIED_STANDARDS_DIMENSIONS,
  };
};

const validateRequiredParts = (archive: DocxArchive, report: MutableReport): boolean => {
  const entries = new Set(archive.entries);
  const missingParts = REQUIRED_PARTS.filter((path) => !entries.has(path));
  if (missingParts.length === 0) {
    report.checks.set("required-parts", "passed");
    return true;
  }

  report.checks.set("required-parts", "failed");
  for (const part of missingParts) {
    addIssue(report, {
      check: "required-parts",
      code: "required-part-missing",
      message: "A required package part is missing.",
      severity: "error",
      part,
    });
  }
  return false;
};

type XmlPartsResult = {
  documentXml: string;
  packageRelationshipsXml: string;
  contentTypesXml: string;
};

const validateXmlParts = async (
  archive: DocxArchive,
  report: MutableReport,
): Promise<XmlPartsResult | null> => {
  const xmlByPath = new Map<string, string>();
  let hasInvalidXml = false;

  for (const part of archive.entries.filter((path) => XML_PART_PATTERN.test(path)).toSorted()) {
    // oxlint-disable-next-line no-await-in-loop -- serialized reads enforce the archive byte budget
    const xml = await archive.readEntryString(part);
    if (xml === null) {
      continue;
    }
    xmlByPath.set(part, xml);

    if (DOCTYPE_PATTERN.test(xml)) {
      hasInvalidXml = true;
      addIssue(report, {
        check: "xml-well-formedness",
        code: "xml-doctype-forbidden",
        message: "XML package parts must not declare a document type.",
        severity: "error",
        part,
      });
      continue;
    }

    if (XMLValidator.validate(xml) !== true) {
      hasInvalidXml = true;
      addIssue(report, {
        check: "xml-well-formedness",
        code: "xml-not-well-formed",
        message: "An XML package part is not well formed.",
        severity: "error",
        part,
      });
    }
  }

  if (hasInvalidXml) {
    report.checks.set("xml-well-formedness", "failed");
    return null;
  }
  report.checks.set("xml-well-formedness", "passed");

  const contentTypesXml = xmlByPath.get("[Content_Types].xml");
  const packageRelationshipsXml = xmlByPath.get("_rels/.rels");
  const documentXml = xmlByPath.get("word/document.xml");
  if (
    contentTypesXml === undefined ||
    packageRelationshipsXml === undefined ||
    documentXml === undefined
  ) {
    report.checks.set("xml-well-formedness", "indeterminate");
    addIssue(report, {
      check: "xml-well-formedness",
      code: "required-xml-unreadable",
      message: "A required XML package part could not be read.",
      severity: "error",
    });
    return null;
  }

  return { contentTypesXml, packageRelationshipsXml, documentXml };
};

type HasRootNamespaceOptions = {
  name: string;
  attributes: Record<string, unknown> | undefined;
  expected: string;
};

const hasRootNamespace = ({ name, attributes, expected }: HasRootNamespaceOptions): boolean => {
  const prefix = getNamespacePrefix(name);
  const attribute = prefix === null ? "xmlns" : `xmlns:${prefix}`;
  return attributes?.[attribute] === expected;
};

const hasDocumentContentType = (contentTypesXml: string): boolean => {
  const root = parseXmlDocument(contentTypesXml);
  if (
    root === null ||
    getLocalName(root.name ?? "") !== "Types" ||
    !hasRootNamespace({
      name: root.name ?? "",
      attributes: root.attributes,
      expected: CONTENT_TYPES_NAMESPACE,
    })
  ) {
    return false;
  }

  return (
    root.elements?.some(
      (element) =>
        element.type === "element" &&
        getLocalName(element.name ?? "") === "Override" &&
        getAttributeAnyPrefix(element, "PartName") === "/word/document.xml" &&
        getAttributeAnyPrefix(element, "ContentType") === DOCUMENT_CONTENT_TYPE,
    ) ?? false
  );
};

const hasMainDocumentRelationship = (packageRelationshipsXml: string): boolean => {
  const root = parseXmlDocument(packageRelationshipsXml);
  if (
    root === null ||
    getLocalName(root.name ?? "") !== "Relationships" ||
    !hasRootNamespace({
      name: root.name ?? "",
      attributes: root.attributes,
      expected: RELATIONSHIPS_NAMESPACE,
    })
  ) {
    return false;
  }

  return (
    root.elements?.some((element) => {
      if (element.type !== "element" || getLocalName(element.name ?? "") !== "Relationship") {
        return false;
      }
      const type = getAttributeAnyPrefix(element, "Type");
      const target = getAttributeAnyPrefix(element, "Target")?.replace(/^\.\//u, "");
      return (
        type?.endsWith("/officeDocument") === true &&
        target === "word/document.xml" &&
        getAttributeAnyPrefix(element, "TargetMode") !== "External"
      );
    }) ?? false
  );
};

const hasDocumentRoot = (documentXml: string): boolean => {
  const root = parseXmlDocument(documentXml);
  return (
    root !== null &&
    getLocalName(root.name ?? "") === "document" &&
    findChildByLocalName(root, "body") !== null
  );
};

const validatePackageRoots = (xml: XmlPartsResult, report: MutableReport): boolean => {
  const invalidParts: string[] = [];
  if (!hasDocumentContentType(xml.contentTypesXml)) {
    invalidParts.push("[Content_Types].xml");
  }
  if (!hasMainDocumentRelationship(xml.packageRelationshipsXml)) {
    invalidParts.push("_rels/.rels");
  }
  if (!hasDocumentRoot(xml.documentXml)) {
    invalidParts.push("word/document.xml");
  }

  if (invalidParts.length === 0) {
    report.checks.set("package-roots", "passed");
    return true;
  }

  report.checks.set("package-roots", "failed");
  for (const part of invalidParts) {
    addIssue(report, {
      check: "package-roots",
      code: "package-root-invalid",
      message:
        "A required package part does not declare the expected root or main document binding.",
      severity: "error",
      part,
    });
  }
  return false;
};

const validateConformanceClass = (documentXml: string, report: MutableReport): void => {
  report.conformanceClass = detectDocxConformanceClass(documentXml);
  if (report.conformanceClass !== DOCX_CONFORMANCE_CLASSES.UNKNOWN) {
    report.checks.set("conformance-class", "passed");
    return;
  }

  report.checks.set("conformance-class", "indeterminate");
  addIssue(report, {
    check: "conformance-class",
    code: "conformance-class-unknown",
    message: "The main document namespace does not identify a supported conformance class.",
    severity: "warning",
    part: "word/document.xml",
  });
};

const findModelValidationError = (error: unknown): DocxModelValidationError | null => {
  let current = error;
  while (current instanceof Error) {
    if (current instanceof DocxModelValidationError) {
      return current;
    }
    current = "cause" in current ? current.cause : undefined;
  }
  return null;
};

const validateCanonicalModel = async (
  bytes: ArrayBuffer | Uint8Array,
  report: MutableReport,
): Promise<void> => {
  try {
    const document = await parseDocx(bytes, {
      detectVariables: false,
      preloadFonts: false,
    });
    const validation = validateFolioDocumentModel(document);
    const warningCount = document.warnings?.length ?? 0;
    if (!validation.valid) {
      report.checks.set("canonical-model", "failed");
    } else if (warningCount > 0) {
      report.checks.set("canonical-model", "indeterminate");
      addIssue(report, {
        check: "canonical-model",
        code: "parser-recovery",
        message: "The canonical parser reported recovery or preservation warnings.",
        severity: "warning",
        count: warningCount,
      });
    } else {
      report.checks.set("canonical-model", "passed");
    }

    for (const issue of validation.issues) {
      addIssue(report, {
        check: "canonical-model",
        code: issue.severity === "error" ? "model-invalid" : "model-warning",
        message: issue.message,
        severity: issue.severity,
        modelPath: issue.path,
      });
    }
  } catch (error) {
    const modelError = findModelValidationError(error);
    if (modelError !== null) {
      report.checks.set("canonical-model", "failed");
      for (const issue of modelError.issues) {
        addIssue(report, {
          check: "canonical-model",
          code: issue.severity === "error" ? "model-invalid" : "model-warning",
          message: issue.message,
          severity: issue.severity,
          modelPath: issue.path,
        });
      }
      return;
    }

    report.checks.set("canonical-model", "indeterminate");
    addIssue(report, {
      check: "canonical-model",
      code: error instanceof DocxParseError ? "parser-unsupported" : "parser-failed",
      message: "The canonical document model could not be assessed.",
      severity: "warning",
    });
  }
};

/** Assess a DOCX package against Folio's named, versioned support profile. */
export const validateDocxConformance = async (
  bytes: ArrayBuffer | Uint8Array,
  options: ValidateDocxConformanceOptions = {},
): Promise<FolioDocxConformanceReport> => {
  const report = createMutableReport();
  const containerType = detectDocxContainerType(bytes);
  if (containerType === DOCX_CONTAINER_TYPES.CFB) {
    report.checks.set("archive-safety", "indeterminate");
    addIssue(report, {
      check: "archive-safety",
      code: "encrypted-container",
      message: "Encrypted containers require decryption before this profile can assess them.",
      severity: "warning",
    });
    return finishReport(report);
  }
  if (containerType !== DOCX_CONTAINER_TYPES.ZIP) {
    report.checks.set("archive-safety", "failed");
    addIssue(report, {
      check: "archive-safety",
      code: "container-not-zip",
      message: "The input is not a supported DOCX package container.",
      severity: "error",
    });
    return finishReport(report);
  }

  let archive: DocxArchive;
  try {
    archive = await loadDocxArchive(bytes, options.archive);
    report.checks.set("archive-safety", "passed");
  } catch (error) {
    report.checks.set("archive-safety", "failed");
    addIssue(report, {
      check: "archive-safety",
      code: error instanceof DocxArchiveError ? `archive-${error.reason}` : "archive-load-failed",
      message: "The DOCX package could not be loaded within the configured safety limits.",
      severity: "error",
    });
    return finishReport(report);
  }

  if (!validateRequiredParts(archive, report)) {
    return finishReport(report);
  }

  let xml: XmlPartsResult | null;
  try {
    xml = await validateXmlParts(archive, report);
  } catch (error) {
    report.checks.set("xml-well-formedness", "indeterminate");
    addIssue(report, {
      check: "xml-well-formedness",
      code: error instanceof DocxArchiveError ? `archive-${error.reason}` : "xml-read-failed",
      message: "The XML package parts could not be read within the configured safety limits.",
      severity: "warning",
    });
    return finishReport(report);
  }
  if (xml === null) {
    return finishReport(report);
  }

  if (!validatePackageRoots(xml, report)) {
    return finishReport(report);
  }

  validateConformanceClass(xml.documentXml, report);
  await validateCanonicalModel(bytes, report);
  return finishReport(report);
};
