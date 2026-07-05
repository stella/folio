// Barrel for the DOCX parse layer. Adapters import numbering helpers from
// `@stll/folio-core/docx`; the parser, serializer, and other modules stay
// reachable at their explicit subpaths via the core `"./*"` export wildcard.
export { getCachedNumberingMap } from "./numberingParser";
