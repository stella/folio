import type { ZipArchive } from "@stll/docx-core/zip";

import { validateDocxPackage } from "../../packages/docx-core/src/validate/docx";

/** Reject malformed test inputs at setup, before a host can edit them. */
export const validateDocxFixture = async (bytes: Uint8Array, name: string) => {
  const result = await validateDocxPackage(bytes);
  if (!result.valid)
    throw new Error(
      `Invalid DOCX fixture ${name}: ${result.code}: ${result.error}`,
    );
  return bytes;
};

/** Shared ZIP producers cannot return or write an unvalidated package. */
export const generateDocxFixture = async (zip: ZipArchive, name: string) =>
  validateDocxFixture(await zip.generate({ level: 9 }), name);
