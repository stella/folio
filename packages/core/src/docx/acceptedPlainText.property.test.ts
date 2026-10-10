import {
  getDocumentText,
  getEndnoteText,
  getFootnoteText,
  getHeaderFooterText,
  getTableText,
  getTextBoxText,
} from "./storyPlainText";
import { expect, test } from "bun:test";
import fc from "fast-check";
import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";
import type { BlockContent, Paragraph, Table } from "../types/document";
import { fromMarkdown, toMarkdown } from "../markdown";
import { blockPlainText } from "./blockPlainText";

const paragraph = (text: string): Paragraph => ({
  type: "paragraph",
  content: [{ type: "run", content: [{ type: "text", text }] }],
});
const cell = (content: BlockContent[]) => ({ type: "tableCell", content }) as const;
const envelope = (content: BlockContent[]): Table => ({
  type: "table",
  rows: [{ type: "tableRow", cells: [cell(content)] }],
});
const markdown = (content: BlockContent[]) => {
  const document = fromMarkdown("");
  document.package.document.content = content;
  return toMarkdown(document, { trackedChanges: "clean", annotations: "strip", comments: "strip" });
};

const READERS = {
  blocks: blockPlainText,
  body: (content: BlockContent[]) => getDocumentText({ content }),
  header: (content: BlockContent[]) =>
    getHeaderFooterText({ type: "header", hdrFtrType: "default", content }),
  footer: (content: BlockContent[]) =>
    getHeaderFooterText({ type: "footer", hdrFtrType: "default", content }),
  footnote: (content: BlockContent[]) => getFootnoteText({ type: "footnote", id: 1, content }),
  endnote: (content: BlockContent[]) => getEndnoteText({ type: "endnote", id: 1, content }),
  table: (content: BlockContent[]) => getTableText(envelope(content)),
  textBox: (content: BlockContent[]) =>
    getTextBoxText({
      type: "textBox",
      size: { width: 914400, height: 457200 },
      content: [envelope(content)],
    }),
} satisfies Record<string, (content: BlockContent[]) => string>;

test(
  "all plain readers resolve generated paragraph and table deletions before projecting text",
  () => {
    assertProperty(
      fc.property(
        fc.array(fc.constantFrom("Článek ", "条項 ", "المادة ", "Clause "), {
          minLength: 1,
          maxLength: 4,
        }),
        fc.integer({ min: 2, max: 4 }),
        fc.constantFrom("del", "moveFrom"),
        (texts, width, kind) => {
          const deletedMark = { kind, info: { id: 91, author: "Test" } };
          const joining = texts.map((text) => ({ ...paragraph(text), pPrMark: deletedMark }));
          const prefix = texts.join("");
          const deletedParagraph = {
            ...paragraph(""),
            pPrMark: deletedMark,
            content: [
              {
                type: kind === "del" ? "deletion" : "moveFrom",
                info: deletedMark.info,
                content: [{ type: "run", content: [{ type: "text", text: prefix }] }],
              },
            ],
          } satisfies Paragraph;
          const columns = Array.from({ length: width }, (_, index) =>
            cell([paragraph(`Cell ${index}.`)]),
          );
          const scenarios = {
            wholeParagraphDeletion: {
              content: [deletedParagraph, paragraph("Tail.")],
              accepted: [paragraph("Tail.")],
              expected: "Tail.",
            },
            finalDeletedParagraph: {
              content: [deletedParagraph],
              accepted: [paragraph("")],
              expected: "",
            },
            finalMarkWithoutDestination: {
              content: [{ ...paragraph(prefix), pPrMark: deletedMark }],
              accepted: [paragraph(prefix)],
              expected: prefix,
            },
            wrappedControl: {
              content: [
                { type: "blockSdt", properties: {}, content: [...joining, paragraph("Tail.")] },
                paragraph("Final."),
              ],
              accepted: [
                { type: "blockSdt", properties: {}, content: [paragraph(`${prefix}Tail.`)] },
                paragraph("Final."),
              ],
              expected: `${prefix}Tail.\nFinal.`,
            },
            wrappedCustomXml: {
              content: [
                {
                  type: "blockCustomXml",
                  openingXml:
                    '<w:customXml xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">',
                  closingXml: "</w:customXml>",
                  content: [...joining, paragraph("Tail.")],
                },
                paragraph("Final."),
              ],
              accepted: [
                {
                  type: "blockCustomXml",
                  openingXml:
                    '<w:customXml xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">',
                  closingXml: "</w:customXml>",
                  content: [paragraph(`${prefix}Tail.`)],
                },
                paragraph("Final."),
              ],
              expected: `${prefix}Tail.\nFinal.`,
            },
            paragraphChain: {
              content: [...joining, paragraph("Tail.")],
              accepted: [paragraph(`${prefix}Tail.`)],
              expected: `${prefix}Tail.`,
            },
            deletedTable: {
              content: [
                ...joining,
                {
                  type: "table",
                  rows: [
                    {
                      type: "tableRow",
                      structuralChange: {
                        type: "tableRowDeletion",
                        info: { id: 92, author: "Test" },
                      },
                      cells: columns,
                    },
                  ],
                },
                paragraph("Tail."),
              ],
              accepted: [paragraph(`${prefix}Tail.`)],
              expected: `${prefix}Tail.`,
            },
            deletedRow: {
              content: [
                {
                  type: "table",
                  rows: [
                    {
                      type: "tableRow",
                      structuralChange: {
                        type: "tableRowDeletion",
                        info: { id: 92, author: "Test" },
                      },
                      cells: columns,
                    },
                    {
                      type: "tableRow",
                      cells: Array.from({ length: width }, (_, index) =>
                        cell([paragraph(`Retained ${index}.`)]),
                      ),
                    },
                  ],
                },
                paragraph("Tail."),
              ],
              accepted: [
                {
                  type: "table",
                  rows: [
                    {
                      type: "tableRow",
                      cells: Array.from({ length: width }, (_, index) =>
                        cell([paragraph(`Retained ${index}.`)]),
                      ),
                    },
                  ],
                },
                paragraph("Tail."),
              ],
              expected: `${Array.from({ length: width }, (_, index) => `Retained ${index}.`).join("\t")}\nTail.`,
            },
            deletedCell: {
              content: [
                {
                  type: "table",
                  rows: [
                    {
                      type: "tableRow",
                      cells: [
                        ...columns.map((value) => ({
                          type: value.type,
                          content: value.content,
                          structuralChange: {
                            type: "tableCellDeletion",
                            info: { id: 93, author: "Test" },
                          },
                        })),
                        cell([paragraph("Retained.")]),
                      ],
                    },
                  ],
                },
                paragraph("Tail."),
              ],
              accepted: [envelope([paragraph("Retained.")]), paragraph("Tail.")],
              expected: "Retained.\nTail.",
            },
            tableCellChain: {
              content: [
                {
                  type: "table",
                  rows: [
                    {
                      type: "tableRow",
                      cells: [cell([...joining, paragraph("Tail.")]), cell([paragraph("Other.")])],
                    },
                  ],
                },
                paragraph("Final."),
              ],
              accepted: [
                {
                  type: "table",
                  rows: [
                    {
                      type: "tableRow",
                      cells: [cell([paragraph(`${prefix}Tail.`)]), cell([paragraph("Other.")])],
                    },
                  ],
                },
                paragraph("Final."),
              ],
              expected: `${prefix}Tail.\tOther.\nFinal.`,
            },
          } satisfies Record<
            string,
            { content: BlockContent[]; accepted: BlockContent[]; expected: string }
          >;
          for (const [scenario, { content, accepted, expected }] of Object.entries(scenarios)) {
            const original = JSON.stringify(content);
            expect(markdown(content), `Markdown / ${scenario}`).toBe(markdown(accepted));
            for (const [reader, read] of Object.entries(READERS)) {
              expect(read(accepted), `${reader} accepted / ${scenario}`).toBe(expected);
              expect(read(content), `${reader} / ${scenario}`).toBe(expected);
              expect(JSON.stringify(content), `${reader} must preserve authored content`).toBe(
                original,
              );
            }
          }
        },
      ),
      {
        numRuns: 24,
        id: "all plain readers resolve generated paragraph and table deletions before projecting text",
      },
    );
  },
  propertyTestTimeout(30_000),
);
