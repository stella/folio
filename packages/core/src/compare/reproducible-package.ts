/**
 * Date comparison output from the comparison's timestamp. Imported entries
 * retain source dates, and saves can stamp `dcterms:modified` from the wall
 * clock; restamping both makes repeated comparisons reproducible.
 */

import { ZipArchive } from "@stll/docx-core/zip";

import { decodeText } from "../utils/textEncoding";

/** ZipArchive deflate level `repackDocx` writes DOCX parts at. */
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
  const zip = await ZipArchive.read(buffer);
  const coreProps = zip.entry(CORE_PROPERTIES_PATH);
  if (coreProps) {
    zip.set(
      CORE_PROPERTIES_PATH,
      withFixedModifiedDate(
        decodeText(await coreProps.bytes(), { ignoreBOM: true }),
        date,
      ),
    );
  }
  return (await zip.generate({ level: DOCX_COMPRESSION_LEVEL, modified: date }))
    .buffer;
};
