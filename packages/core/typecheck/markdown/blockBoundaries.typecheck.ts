import type { BlockContent, Paragraph } from "../../src/types/document";

const paragraph = (text: string): Paragraph => ({
  type: "paragraph",
  content: [{ type: "run", content: [{ type: "text", text }] }],
});

type BoundaryFixtures = {
  [Kind in BlockContent["type"]]: {
    block: Extract<BlockContent, { type: Kind }>;
    marker: string | null;
  };
};

export const MARKDOWN_BLOCK_BOUNDARY_FIXTURES = {
  paragraph: { block: paragraph("MIDDLE"), marker: "MIDDLE" },
  table: {
    block: {
      type: "table",
      rows: [
        {
          type: "tableRow",
          cells: [{ type: "tableCell", content: [paragraph("MIDDLE")] }],
        },
      ],
    },
    marker: "MIDDLE",
  },
  blockSdt: {
    block: {
      type: "blockSdt",
      properties: { sdtType: "richText" },
      content: [paragraph("MIDDLE")],
    },
    marker: "MIDDLE",
  },
  blockCustomXml: {
    block: {
      type: "blockCustomXml",
      openingXml:
        '<w:customXml xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">',
      closingXml: "</w:customXml>",
      content: [paragraph("MIDDLE")],
    },
    marker: "MIDDLE",
  },
  preservedBlock: {
    block: {
      type: "preservedBlock",
      xml: '<w:altChunk xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>',
      readerText: "MIDDLE",
    },
    marker: "MIDDLE",
  },
  bookmarkStart: {
    block: { type: "bookmarkStart", id: 1, name: "boundary" },
    marker: null,
  },
  bookmarkEnd: { block: { type: "bookmarkEnd", id: 1 }, marker: null },
} as const satisfies BoundaryFixtures;

// @ts-expect-error a boundary census missing any block kind must fail
const incompleteCensus: BoundaryFixtures = {
  paragraph: MARKDOWN_BLOCK_BOUNDARY_FIXTURES.paragraph,
};
void incompleteCensus;
