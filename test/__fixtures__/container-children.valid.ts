/**
 * The same container through the shared dispatcher. The handler map is total
 * over the children the schema declares, and anything else reaches the ordered
 * verbatim sink instead of the floor.
 */

import { CAPTURE, dispatchChildren } from "../../packages/core/src/docx/containerChildren";
import type { XmlElement } from "../../packages/core/src/docx/xmlParser";

export const parseThing = (container: XmlElement) => {
  const modelled: string[] = [];
  const preserved = dispatchChildren({
    element: container,
    container: "w:comment",
    modelledCount: () => modelled.length,
    handlers: {
      p: () => modelled.push("paragraph"),
      altChunk: CAPTURE,
      bookmarkEnd: CAPTURE,
      bookmarkStart: CAPTURE,
      commentRangeEnd: CAPTURE,
      commentRangeStart: CAPTURE,
      customXml: CAPTURE,
      customXmlDelRangeEnd: CAPTURE,
      customXmlDelRangeStart: CAPTURE,
      customXmlInsRangeEnd: CAPTURE,
      customXmlInsRangeStart: CAPTURE,
      customXmlMoveFromRangeEnd: CAPTURE,
      customXmlMoveFromRangeStart: CAPTURE,
      customXmlMoveToRangeEnd: CAPTURE,
      customXmlMoveToRangeStart: CAPTURE,
      del: CAPTURE,
      ins: CAPTURE,
      moveFrom: CAPTURE,
      moveFromRangeEnd: CAPTURE,
      moveFromRangeStart: CAPTURE,
      moveTo: CAPTURE,
      moveToRangeEnd: CAPTURE,
      moveToRangeStart: CAPTURE,
      permEnd: CAPTURE,
      permStart: CAPTURE,
      proofErr: CAPTURE,
      sdt: CAPTURE,
      tbl: CAPTURE,
    },
  });
  return { modelled, preserved };
};
