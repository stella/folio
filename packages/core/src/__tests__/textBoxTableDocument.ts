import { parseDocx } from "../docx/parser";
import { createEmptyDocx, repackDocx } from "../docx/rezip";
import type { Document, Shape } from "../types/document";

export const buildTextBoxTableDocument = async (cellText = "Cell value"): Promise<ArrayBuffer> => {
  const source = await createEmptyDocx();
  const document = await parseDocx(source, { detectVariables: false, preloadFonts: false });
  document.package.document.content = [
    {
      type: "paragraph",
      paraId: "A2000001",
      content: [
        {
          type: "run",
          content: [
            {
              type: "shape",
              shape: {
                type: "shape",
                shapeType: "textBox",
                id: "42",
                size: { width: 1_828_800, height: 914_400 },
                textBody: {
                  margins: { top: 45_720, bottom: 45_720, left: 91_440, right: 91_440 },
                  content: [
                    {
                      type: "paragraph",
                      paraId: "A2000002",
                      content: [{ type: "run", content: [{ type: "text", text: "Before table" }] }],
                    },
                    {
                      type: "table",
                      rows: [
                        {
                          type: "tableRow",
                          cells: [
                            {
                              type: "tableCell",
                              content: [
                                {
                                  type: "paragraph",
                                  paraId: "A2000003",
                                  content: [
                                    {
                                      type: "run",
                                      content: [{ type: "text", text: cellText }],
                                    },
                                  ],
                                },
                              ],
                            },
                          ],
                        },
                      ],
                    },
                    {
                      type: "paragraph",
                      paraId: "A2000004",
                      content: [{ type: "run", content: [{ type: "text", text: "After table" }] }],
                    },
                  ],
                },
              },
            },
          ],
        },
      ],
    },
  ];

  return repackDocx(document, { updateModifiedDate: false });
};

export const findTextBoxShape = (document: Document): Shape => {
  for (const block of document.package.document.content) {
    if (block.type !== "paragraph") {
      continue;
    }
    for (const content of block.content) {
      if (content.type !== "run") {
        continue;
      }
      for (const runContent of content.content) {
        if (runContent.type === "shape" && runContent.shape.textBody) {
          return runContent.shape;
        }
      }
    }
  }
  throw new Error("expected a shape with text content");
};
