/**
 * The last clock in the compare path is the ZIP container itself.
 *
 * Every part the serializer rewrites is stored with JSZip's default entry
 * date, which is `new Date()`. The XML is identical between two runs, but the
 * DOS timestamp fields in the local headers are not, and they only agree when
 * both runs land in the same two-second bucket — so the packages match most of
 * the time and differ occasionally, which is worse than differing always.
 *
 * Restamping every entry from the comparison's own timestamp removes it. It
 * also states the truth about the package: a generated redline is dated by the
 * comparison that produced it, not by the second it happened to be written.
 */

import JSZip from "jszip";

/** JSZip deflate level `repackDocx` writes DOCX parts at. */
const DOCX_COMPRESSION_LEVEL = 6;

export const withFixedPackageDates = async (
  buffer: ArrayBuffer,
  date: Date,
): Promise<ArrayBuffer> => {
  const zip = await JSZip.loadAsync(buffer);
  zip.forEach((_path, file) => {
    file.date = date;
  });
  return await zip.generateAsync({
    type: "arraybuffer",
    compression: "DEFLATE",
    compressionOptions: { level: DOCX_COMPRESSION_LEVEL },
  });
};
