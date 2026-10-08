import { writeZipPart } from "@stll/docx-core/zip";
/**
 * Date comparison output from the comparison's timestamp. Imported entries
 * retain source dates, and saves can stamp `dcterms:modified` from the wall
 * clock; restamping both makes repeated comparisons reproducible.
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
    writeZipPart({
      zip,
      path: CORE_PROPERTIES_PATH,
      data: withFixedModifiedDate(await coreProps.async("text"), date),
      options: {
        compression: "DEFLATE",
        compressionOptions: { level: DOCX_COMPRESSION_LEVEL },
      },
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
