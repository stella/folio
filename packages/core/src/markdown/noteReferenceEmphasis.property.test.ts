import { expect, test } from "bun:test";
import fc from "fast-check";

import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";
import type { Document, Run } from "../types/document";
import { toMarkdown } from "./index";

const text = (value: string): Run => ({ type: "run", content: [{ type: "text", text: value }] });
const edge = fc
  .array(fc.constantFrom("a", "é", "😀", " ", ".", ")", "*", "_", "~"), { maxLength: 8 })
  .map((chars) => chars.join(""));

// The emphasis generator previously emitted text only and stripped notes.
// Note-reference syntax must remain indivisible at every emphasis boundary.
test(
  "every emitted note reference survives all emphasis combinations and neighboring text",
  async () => {
    let cases = 0;
    await assertProperty(
      fc.property(fc.record({ before: edge, after: edge }), ({ before, after }) => {
        for (const kind of ["footnote", "endnote"] as const) {
          for (let mask = 0; mask < 8; mask++) {
            const marker = kind === "footnote" ? "[^1]" : "[^e1]";
            const document: Document = {
              package: {
                document: {
                  content: [
                    {
                      type: "paragraph",
                      content: [
                        text(`Start ${before}`),
                        {
                          type: "run",
                          formatting: {
                            bold: Boolean(mask & 1),
                            italic: Boolean(mask & 2),
                            strike: Boolean(mask & 4),
                          },
                          content: [
                            { type: kind === "footnote" ? "footnoteRef" : "endnoteRef", id: 17 },
                          ],
                        },
                        text(`${after} end`),
                      ],
                    },
                  ],
                },
                ...(kind === "footnote"
                  ? {
                      footnotes: [
                        {
                          type: "footnote",
                          id: 17,
                          content: [{ type: "paragraph", content: [text("Note text.")] }],
                        },
                      ],
                    }
                  : {
                      endnotes: [
                        {
                          type: "endnote",
                          id: 17,
                          content: [{ type: "paragraph", content: [text("Note text.")] }],
                        },
                      ],
                    }),
              },
            };
            const markdown = toMarkdown(document, {
              annotations: "strip",
              trackedChanges: "clean",
              comments: "strip",
              footnotes: "keep",
            });
            const boundary = markdown.lastIndexOf("\n\n");
            const body = markdown.slice(0, boundary);
            expect(body.match(/(?<!\\)\[\^e?\d+\]/gu)).toEqual([marker]);
            expect(markdown.slice(boundary + 2)).toBe(`${marker}: Note text.`);
            cases++;
          }
        }
      }),
      {
        numRuns: 50,
        id: "every emitted note reference survives all emphasis combinations and neighboring text",
      },
    );
    expect(cases).toBeGreaterThan(0);
  },
  propertyTestTimeout(30_000),
);
