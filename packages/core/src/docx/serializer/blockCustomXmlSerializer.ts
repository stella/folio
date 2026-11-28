import { normalizeRevisionId } from "@stll/docx-core/model";

import type {
  BlockContent,
  BlockCustomXml,
  CustomXmlMarkupRevision,
} from "../../types/document";
import { serializeTrackedChangeAttributes } from "./trackedChangeAttributes";

const startTag = (name: string, record: CustomXmlMarkupRevision): string =>
  `<w:${name} ${serializeTrackedChangeAttributes(record.info)}/>`;

const endTag = (name: string, record: CustomXmlMarkupRevision): string =>
  `<w:${name} w:id="${String(normalizeRevisionId(record.info.id))}"/>`;

/**
 * The range markers written around a block custom XML element for its own
 * records: the insertion outermost, as a deletion of tags that were inserted
 * comes after their insertion.
 */
export const customXmlMarkupRevisionXml = (
  block: Pick<BlockCustomXml, "insertion" | "deletion">,
): { before: string; after: string } => {
  let before = "";
  let after = "";
  if (block.insertion !== undefined) {
    before += startTag("customXmlInsRangeStart", block.insertion);
    after = endTag("customXmlInsRangeEnd", block.insertion) + after;
  }
  if (block.deletion !== undefined) {
    before += startTag("customXmlDelRangeStart", block.deletion);
    after = endTag("customXmlDelRangeEnd", block.deletion) + after;
  }
  return { before, after };
};

/**
 * The element with the range markers of its own records around it
 * (ECMA-376 Part 1, §17.13.5.4–7).
 */
export const serializeBlockCustomXml = (
  block: BlockCustomXml,
  serializeChild: (child: BlockContent) => string,
): string => {
  const records = customXmlMarkupRevisionXml(block);
  return (
    records.before +
    block.openingXml +
    block.content.map(serializeChild).join("") +
    block.closingXml +
    records.after
  );
};
