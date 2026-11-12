import type {
  MediaFile,
  RelationshipMap,
  TableCellBlock,
  Theme,
} from "../types/document";
import type { DeepReadonly } from "../types/documentModel";
import { blockCustomXmlShell } from "./blockCustomXmlShell";
import { parseBookmarkEnd, parseBookmarkStart } from "./bookmarkParser";
import type { NumberingMap } from "./numberingParser";
import type { PreviewLedger } from "./previewBudget";
import { captureSdtSiblingMarkers, parseSdtProperties } from "./sdtProperties";
import type { StyleMap } from "./styleParser";
import type { ParagraphParserFn, TableParserFn } from "./textBoxParser";
import { captureVerbatimXml } from "./verbatimCapture";
import {
  findChildByLocalName,
  getChildElements,
  WORDPROCESSINGML_NAMESPACE_URIS,
  type XmlElement,
} from "./xmlParser";

/**
 * Parse text box content with provided parser functions
 * This avoids circular dependencies by accepting parser functions as parameters
 */
export function parseTextBoxContent(
  txbxContent: XmlElement | null,
  parseParagraph: ParagraphParserFn,
  parseTable: TableParserFn | null,
  styles: StyleMap | null,
  theme: DeepReadonly<Theme> | null,
  numbering: NumberingMap | null,
  rels: RelationshipMap | null,
  media: Map<string, MediaFile> | null,
  previews: PreviewLedger,
): TableCellBlock[] {
  if (!txbxContent) {
    return [];
  }

  const readBlocks = (container: XmlElement): TableCellBlock[] => {
    const blocks: TableCellBlock[] = [];
    for (const child of getChildElements(container)) {
      const name = child.name || "";
      const colonIdx = name.indexOf(":");
      const localName = colonIdx !== -1 ? name.slice(colonIdx + 1) : name;
      const isWordElement = WORDPROCESSINGML_NAMESPACE_URIS.has(
        child.namespaceUri ?? "",
      );

      if (localName === "p") {
        blocks.push(
          parseParagraph(child, styles, theme, numbering, rels, media, {
            previews,
          }),
        );
      } else if (localName === "tbl" && parseTable) {
        const table = parseTable(child, styles, theme, numbering, rels, media, {
          previews,
        });
        if (table) {
          blocks.push(table);
        }
      } else if (localName === "sdt" && isWordElement) {
        const properties = parseSdtProperties(
          findChildByLocalName(child, "sdtPr"),
          findChildByLocalName(child, "sdtEndPr"),
        );
        const captured = captureSdtSiblingMarkers(child);
        if (captured.before.length > 0) {
          properties.rawSdtChildrenBeforeContent = captured.before;
        }
        if (captured.after.length > 0) {
          properties.rawSdtChildrenAfterContent = captured.after;
        }
        const sdtContent = findChildByLocalName(child, "sdtContent");
        blocks.push({
          type: "blockSdt",
          properties,
          content: sdtContent ? readBlocks(sdtContent) : [],
        });
      } else if (localName === "customXml" && isWordElement) {
        blocks.push({
          type: "blockCustomXml",
          ...blockCustomXmlShell(child),
          content: readBlocks(child),
        });
      } else if (localName === "bookmarkStart" && isWordElement) {
        blocks.push(parseBookmarkStart(child));
      } else if (localName === "bookmarkEnd" && isWordElement) {
        blocks.push(parseBookmarkEnd(child));
      } else {
        // Whatever else the box holds stays where it was written: a rewrite of
        // the paragraph that hosts the box writes this list back as the box.
        blocks.push({ type: "preservedBlock", xml: captureVerbatimXml(child) });
      }
    }
    return blocks;
  };

  return readBlocks(txbxContent);
}
