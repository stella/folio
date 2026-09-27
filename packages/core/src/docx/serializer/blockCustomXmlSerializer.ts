import type { BlockContent, BlockCustomXml } from "../../types/document";

export const serializeBlockCustomXml = (
  block: BlockCustomXml,
  serializeChild: (child: BlockContent) => string,
): string => block.openingXml + block.content.map(serializeChild).join("") + block.closingXml;
