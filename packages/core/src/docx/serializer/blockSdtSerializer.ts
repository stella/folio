/**
 * Block-level SDT serializer.
 *
 * The property set comes from the one writer, `serializeSdtProperties`, which
 * merges the modelled children with the ones the parser kept as bytes;
 * `<w:sdtEndPr>` is still replayed from `rawEndPropertiesXml`.
 *
 * Sharing the helper between the document body and the header/footer
 * serializers keeps body↔HF parity in one place.
 */

import type { BlockContent, BlockSdt } from "../../types/document";
import { serializeSdtProperties } from "./sdtPropertiesSerializer";
import { isSingleWellFormedElement } from "./xmlUtils";

export function serializeBlockSdt(
  blockSdt: BlockSdt,
  serializeChild: (block: BlockContent) => string,
): string {
  const props = blockSdt.properties;
  const sdtPrXml = serializeSdtProperties(props);
  const sdtEndPrXml =
    props.rawEndPropertiesXml && isSingleWellFormedElement(props.rawEndPropertiesXml, "sdtEndPr")
      ? props.rawEndPropertiesXml
      : "";
  const contentXml = blockSdt.content.map(serializeChild).join("");
  // Replay any direct sdt children that lived OUTSIDE sdtContent at parse
  // time (range markers per MS-OE376 §2.5.2.30: bookmark / comment /
  // tracked-change / custom XML ranges that span an SDT boundary). Position
  // matters — the captured before/after strings preserve which side of
  // sdtContent each marker sat on.
  const before = props.rawSdtChildrenBeforeContent ?? "";
  const after = props.rawSdtChildrenAfterContent ?? "";
  return `<w:sdt>${sdtPrXml}${sdtEndPrXml}${before}<w:sdtContent>${contentXml}</w:sdtContent>${after}</w:sdt>`;
}
