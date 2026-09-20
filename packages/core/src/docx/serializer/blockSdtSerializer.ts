/**
 * Block-level SDT serializer.
 *
 * The property elements are `sdtPropertiesSerializer`'s, shared with the
 * inline, row-level and cell-level controls so the rule for replaying or
 * rebuilding them is written once.
 *
 * Sharing this helper between the document body and the header/footer
 * serializers keeps body↔HF parity in one place.
 */

import type { BlockContent, BlockSdt } from "../../types/document";
import { serializeSdtPropertyElements } from "./sdtPropertiesSerializer";

export function serializeBlockSdt(
  blockSdt: BlockSdt,
  serializeChild: (block: BlockContent) => string,
): string {
  const props = blockSdt.properties;
  const propertyElements = serializeSdtPropertyElements(props);
  const contentXml = blockSdt.content.map(serializeChild).join("");
  // Replay any direct sdt children that lived OUTSIDE sdtContent at parse
  // time (range markers per MS-OE376 §2.5.2.30: bookmark / comment /
  // tracked-change / custom XML ranges that span an SDT boundary). Position
  // matters — the captured before/after strings preserve which side of
  // sdtContent each marker sat on.
  const before = props.rawSdtChildrenBeforeContent ?? "";
  const after = props.rawSdtChildrenAfterContent ?? "";
  return `<w:sdt>${propertyElements}${before}<w:sdtContent>${contentXml}</w:sdtContent>${after}</w:sdt>`;
}
