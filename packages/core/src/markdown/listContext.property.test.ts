import { expect, test } from "bun:test";
import fc from "fast-check";

import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";
import type { BlockContent, ListRendering, Paragraph } from "../types/document";
import { toMarkdown } from "./index";

const paragraph = (text: string, listRendering?: ListRendering): Paragraph => ({
  type: "paragraph",
  content: [{ type: "run", content: [{ type: "text", text }] }],
  ...(listRendering === undefined ? {} : { listRendering }),
});

// Markdown's ancestry is emitted list items, not every logical OOXML level.
// The bridge generator previously only produced complete, adjacent list trees.
test(
  "orphan list levels never inherit absent or interrupted Markdown ancestors",
  async () => {
    let cases = 0;
    await assertProperty(
      fc.property(fc.integer({ min: 1, max: 999 }), (start) => {
        for (let level = 0; level <= 8; level++) {
          const cellMarkdown = toMarkdown({
            package: {
              document: {
                content: [
                  {
                    type: "table",
                    rows: [
                      {
                        type: "tableRow",
                        cells: [
                          {
                            type: "tableCell",
                            content: [
                              paragraph("Parent", {
                                marker: "-",
                                level: 0,
                                numId: 1,
                                isBullet: true,
                              }),
                            ],
                          },
                          {
                            type: "tableCell",
                            content: [
                              paragraph("Child", { marker: "-", level, numId: 2, isBullet: true }),
                            ],
                          },
                        ],
                      },
                    ],
                  },
                ],
              },
            },
          });
          expect(cellMarkdown.split("\n").at(0)).toBe("| - Parent | - Child |");
          for (const marker of ["-", "i.", "1.2.", `${start}.`]) {
            const list = { marker, level, numId: 2, isBullet: marker === "-" };
            for (const boundary of [
              "start",
              "prose",
              "heading",
              "hidden",
              "custom",
              "table",
              "control",
              "preserved",
            ] as const) {
              const blocks: BlockContent[] = [];
              if (boundary !== "start") {
                blocks.push(
                  paragraph("Parent", { marker: `${start}.`, level: 0, numId: 1, isBullet: false }),
                );
              }
              switch (boundary) {
                case "start":
                  break;
                case "prose":
                  blocks.push(paragraph("Boundary"));
                  break;
                case "heading":
                  blocks.push({
                    ...paragraph("Boundary"),
                    formatting: { outlineLevel: { kind: "heading", level: 0 } },
                  });
                  break;
                case "hidden":
                  blocks.push(paragraph("Boundary", { ...list, markerHidden: true }));
                  break;
                case "custom":
                  blocks.push(
                    paragraph("Boundary", { marker: "a)", level: 0, numId: 3, isBullet: false }),
                  );
                  break;
                case "table":
                  blocks.push({
                    type: "table",
                    rows: [
                      {
                        type: "tableRow",
                        cells: [
                          {
                            type: "tableCell",
                            content: [
                              paragraph("Cell", {
                                marker: "-",
                                level: 0,
                                numId: 3,
                                isBullet: true,
                              }),
                            ],
                          },
                        ],
                      },
                    ],
                  });
                  break;
                case "control":
                  blocks.push({
                    type: "blockSdt",
                    properties: {},
                    content: [
                      paragraph("Boundary", { marker: "-", level: 0, numId: 3, isBullet: true }),
                    ],
                  });
                  break;
                case "preserved":
                  blocks.push({
                    type: "preservedBlock",
                    xml: '<w:altChunk xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>',
                    readerText: "Boundary",
                  });
                  break;
                default: {
                  const exhaustive: never = boundary;
                  throw exhaustive;
                }
              }
              blocks.push(paragraph("Target", list));
              const markdown = toMarkdown({ package: { document: { content: blocks } } });
              expect(markdown.split("\n").at(-1)).toBe(`${marker} Target`);
              cases++;
            }
          }
        }
      }),
      { numRuns: 10 },
    );
    expect(cases).toBeGreaterThan(0);
  },
  propertyTestTimeout(30_000),
);

test(
  "skipped logical levels nest only under emitted native list ancestors",
  async () => {
    let cases = 0;
    await assertProperty(
      fc.property(fc.integer({ min: 1, max: 999 }), (start) => {
        for (let parentLevel = 0; parentLevel < 8; parentLevel++) {
          for (let childLevel = parentLevel + 1; childLevel <= 8; childLevel++) {
            const parent = { marker: `${start}.`, level: parentLevel, numId: 1, isBullet: false };
            const child = { marker: "-", level: childLevel, numId: 2, isBullet: true };
            const markdown = toMarkdown({
              package: {
                document: {
                  content: [
                    paragraph("Parent", parent),
                    paragraph("Child", child),
                    paragraph("Sibling", parent),
                  ],
                },
              },
            });
            expect(markdown).toBe(
              `${start}. Parent\n${" ".repeat(String(start).length + 2)}- Child\n${start}. Sibling`,
            );
            cases++;
          }
        }
      }),
      { numRuns: 10 },
    );
    expect(cases).toBeGreaterThan(0);
  },
  propertyTestTimeout(30_000),
);
