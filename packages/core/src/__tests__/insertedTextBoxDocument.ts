import type { Paragraph } from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";

export const INSERTED_TEXT_BOX_STYLE = "InsertedTextBoxStyle";

type InsertedTextBoxDocumentOptions = {
  side: "base" | "revised";
  innerContent: "paragraph" | "tableCell";
  suffix: string;
  innerCount?: number;
  carrierCount?: number;
};

export const insertedTextBoxDocument = ({
  side,
  innerContent,
  suffix,
  innerCount = 1,
  carrierCount = 1,
}: InsertedTextBoxDocumentOptions) => {
  const document = createEmptyDocument();
  document.package.styles = {
    styles: [
      { type: "paragraph", styleId: "Normal", name: "Normal", default: true },
      {
        type: "paragraph",
        styleId: INSERTED_TEXT_BOX_STYLE,
        name: "Inserted text box style",
        rPr: { bold: side === "revised" },
      },
    ],
  };
  document.package.document.content = [
    {
      type: "paragraph",
      paraId: "23000001",
      content: [{ type: "run", content: [{ type: "text", text: "Unchanged anchor" }] }],
    },
    ...(side === "revised"
      ? Array.from({ length: carrierCount }, (_, carrierIndex) => {
          const carrierSuffix =
            carrierCount === 1 ? suffix : `${suffix} carrier ${carrierIndex + 1}`;
          const innerParagraphs = Array.from(
            { length: innerCount },
            (_innerValue, innerIndex) =>
              ({
                type: "paragraph",
                paraId: (0x23000003 + carrierIndex * 16 + innerIndex).toString(16).toUpperCase(),
                formatting: { styleId: INSERTED_TEXT_BOX_STYLE },
                content: [
                  {
                    type: "run",
                    content: [
                      {
                        type: "text",
                        text: `Inner text ${carrierSuffix}${innerCount === 1 ? "" : ` ${innerIndex + 1}`}`,
                      },
                    ],
                  },
                ],
              }) as const satisfies Paragraph,
          );
          return {
            type: "paragraph" as const,
            paraId: (0x23000002 + carrierIndex * 16).toString(16).toUpperCase(),
            formatting: { styleId: INSERTED_TEXT_BOX_STYLE },
            content: [
              {
                type: "run" as const,
                content: [
                  { type: "text" as const, text: `Parent text ${carrierSuffix}` },
                  {
                    type: "shape" as const,
                    shape: {
                      type: "shape" as const,
                      shapeType: "textBox" as const,
                      id: String(43 + carrierIndex),
                      size: { width: 1_828_800, height: 914_400 },
                      textBody: {
                        content:
                          innerContent === "paragraph"
                            ? innerParagraphs
                            : [
                                {
                                  type: "table" as const,
                                  rows: [
                                    {
                                      type: "tableRow" as const,
                                      cells: [
                                        { type: "tableCell" as const, content: innerParagraphs },
                                      ],
                                    },
                                  ],
                                },
                              ],
                      },
                    },
                  },
                ],
              },
            ],
          };
        })
      : []),
  ];
  return document;
};
