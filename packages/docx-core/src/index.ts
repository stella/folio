export type {
  BlockContent,
  BreakContent,
  Document,
  DocumentBody,
  DocxConformanceClass,
  DocxPackage,
  Paragraph,
  ParagraphContent,
  Run,
  RunContent,
  SectionProperties,
  Style,
  Table,
  TableCell,
  TableRow,
  TextContent,
} from "./model/document";
export { DOCX_CONFORMANCE_CLASSES } from "./model/document";
export {
  compileLegalSourceToDocument,
  compileLegalSourceToDocx,
  parseLegalSource,
  validateLegalDraft,
} from "./legal-source";
export type {
  Autofix,
  CompiledLegalDocument,
  LegalDraft,
  LegalDraftBlock,
  LegalDraftDiagnostic,
  LegalSourceCompileOptions,
  LegalSourceCompileResult,
  LegalSourceDocxCompileResult,
  LegalSourceParseResult,
} from "./legal-source";
export { serializeDocumentToDocx } from "./serialize/docx";
export {
  assertValidDocumentModel,
  validateDocxPackage,
  validateDocumentModel,
} from "./validate/docx";
export type {
  ValidateDocumentModelIssue,
  ValidateDocumentModelResult,
  ValidateDocxPackageResult,
} from "./validate/docx";
