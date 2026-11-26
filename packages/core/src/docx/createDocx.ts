import type { Document } from "../types/document";
import type { DocumentPropertiesOptions } from "./rezip";

/** Create a DOCX from a document, loading the package writer on demand. */
export async function createDocx(
  document: Document,
  properties?: DocumentPropertiesOptions,
): Promise<ArrayBuffer> {
  const writer = await import("./rezip");
  return writer.createDocx(document, properties);
}
