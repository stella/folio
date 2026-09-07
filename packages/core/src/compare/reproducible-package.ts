/**
 * The last two clocks in the compare path are outside the document body.
 *
 * Every part the serializer rewrites is stored with JSZip's default entry
 * date, which is `new Date()`. The XML is identical between two runs, but the
 * DOS timestamp fields in the local headers are not, and they only agree when
 * both runs land in the same two-second bucket — so the packages match most of
 * the time and differ occasionally, which is worse than differing always.
 *
 * The save also stamps `dcterms:modified` in `docProps/core.xml` from the wall
 * clock, which is the same failure one part deeper: two runs over identical
 * inputs differ in that part alone.
 *
 * Restamping both from the comparison's own timestamp removes them. It also
 * states the truth about the package: a generated redline is dated by the
 * comparison that produced it, not by the second it happened to be written.
 */

import JSZip from "jszip";

/** JSZip deflate level `repackDocx` writes DOCX parts at. */
const DOCX_COMPRESSION_LEVEL = 6;

const CORE_PROPERTIES_PATH = "docProps/core.xml";

const MODIFIED_ELEMENT = /<dcterms:modified[^<>]*>[^<]*<\/dcterms:modified>/u;

/**
 * Rewrite `dcterms:modified` where the save already wrote one. An absent
 * element stays absent: the comparison edits the document it was handed, and
 * synthesizing metadata the input never carried is a different decision.
 */
const withFixedModifiedDate = (corePropsXml: string, date: Date): string =>
  corePropsXml.replace(
    MODIFIED_ELEMENT,
    `<dcterms:modified xsi:type="dcterms:W3CDTF">${date.toISOString()}</dcterms:modified>`,
  );

export const withFixedPackageDates = async (
  buffer: ArrayBuffer,
  date: Date,
): Promise<ArrayBuffer> => {
  const zip = await JSZip.loadAsync(buffer);
  const coreProps = zip.file(CORE_PROPERTIES_PATH);
  if (coreProps) {
    zip.file(CORE_PROPERTIES_PATH, withFixedModifiedDate(await coreProps.async("text"), date), {
      compression: "DEFLATE",
      compressionOptions: { level: DOCX_COMPRESSION_LEVEL },
    });
  }
  zip.forEach((_path, file) => {
    file.date = date;
  });
  return await zip.generateAsync({
    type: "arraybuffer",
    compression: "DEFLATE",
    compressionOptions: { level: DOCX_COMPRESSION_LEVEL },
  });
};
