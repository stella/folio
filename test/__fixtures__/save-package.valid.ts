/** A save entry that reads the package it writes over through the owner. */

import { readPackageForSave } from "../../packages/core/src/docx/mainDocumentPart";

export const savedPaths = async (originalBuffer: ArrayBuffer): Promise<string[]> =>
  (await readPackageForSave(originalBuffer)).paths();
