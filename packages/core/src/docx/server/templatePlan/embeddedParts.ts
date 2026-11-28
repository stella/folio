/**
 * The parts an object may carry whole from the source into the output:
 * pictures, and embedded objects (an OLE object's storage, an embedded
 * package such as a workbook). Each is one part with no relationships of
 * its own, so it travels as bytes under a new name.
 */
import { RELATIONSHIP_TYPES } from "../../relsParser";

/** An embedded package (a workbook, a document, a presentation). */
export const PACKAGE_RELATIONSHIP_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package";

/** Relationship types whose target travels whole with its object. */
export const EMBEDDED_RELATIONSHIP_TYPES: ReadonlySet<string> = new Set([
  RELATIONSHIP_TYPES.image,
  RELATIONSHIP_TYPES.oleObject,
  PACKAGE_RELATIONSHIP_TYPE,
]);

const PICTURE_CONTENT_TYPES: Readonly<Record<string, string>> = {
  png: "image/png",
  jpg: "image/jpeg",
  jpeg: "image/jpeg",
  gif: "image/gif",
  bmp: "image/bmp",
  tif: "image/tiff",
  tiff: "image/tiff",
  emf: "image/x-emf",
  wmf: "image/x-wmf",
  svg: "image/svg+xml",
  webp: "image/webp",
};

const OFFICE = "application/vnd.openxmlformats-officedocument.";

const EMBEDDING_CONTENT_TYPES: Readonly<Record<string, string>> = {
  bin: `${OFFICE}oleObject`,
  xls: "application/vnd.ms-excel",
  doc: "application/msword",
  ppt: "application/vnd.ms-powerpoint",
  xlsx: `${OFFICE}spreadsheetml.sheet`,
  xlsm: "application/vnd.ms-excel.sheet.macroEnabled.12",
  xlsb: "application/vnd.ms-excel.sheet.binary.macroEnabled.12",
  docx: `${OFFICE}wordprocessingml.document`,
  docm: "application/vnd.ms-word.document.macroEnabled.12",
  pptx: `${OFFICE}presentationml.presentation`,
  pptm: "application/vnd.ms-powerpoint.presentation.macroEnabled.12",
  sldx: `${OFFICE}presentationml.slide`,
  vsdx: "application/vnd.ms-visio.drawing",
};

/** Whether a part of `relationshipType` is a picture (else an embedding). */
export const isPicturePart = (relationshipType: string): boolean =>
  relationshipType === RELATIONSHIP_TYPES.image;

/**
 * The content type the output declares for a carried part, or undefined
 * when its extension is not one the output can declare for that kind.
 */
export const embeddedContentType = (
  relationshipType: string,
  extension: string,
): string | undefined => {
  if (isPicturePart(relationshipType)) {
    return PICTURE_CONTENT_TYPES[extension];
  }
  return EMBEDDED_RELATIONSHIP_TYPES.has(relationshipType)
    ? EMBEDDING_CONTENT_TYPES[extension]
    : undefined;
};

/** The extension of a package path, lower case; empty without one. */
export const extensionOf = (path: string): string => {
  const slash = path.lastIndexOf("/");
  const dot = path.lastIndexOf(".");
  return dot <= slash ? "" : path.slice(dot + 1).toLowerCase();
};
