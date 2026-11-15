/** Deliberate violation: a save entry that reads the package it writes over itself. */

import { ZipArchive } from "@stll/docx-core/zip";

export const savedPaths = async (originalBuffer: ArrayBuffer): Promise<string[]> =>
  (await ZipArchive.read(originalBuffer)).paths();
