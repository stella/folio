import { expect, test } from "bun:test";

import { prepareCanonicalDocxInput } from "../docx/canonicalSessionInput";
import { parseDocx } from "../docx/parser";
import { createDocx } from "../docx/rezip";
import { headerFooterToProseDoc } from "../prosemirror/conversion/toProseDoc";
import type { BlockContent, Document, Paragraph } from "../types/document";
import { readModel } from "../types/documentModel";
import { createCanonicalSession } from "./canonicalSession";

const item = (text: string, ilvl: number, style = false): Paragraph => ({
  type: "paragraph",
  formatting: style
    ? { styleId: "ListStyle" }
    : { numPr: { kind: "reference", numId: 1, ilvl } },
  content: [{ type: "run", content: [{ type: "text", text }] }],
});

const storyContent = (style: boolean): BlockContent[] => [
  item("A", 0, style),
  item("B", 1, style),
  item("C", 0, style),
];

const documentWith = (style: boolean): Document => {
  const page = {
    pageWidth: 12240,
    pageHeight: 15840,
    headerReferences: [{ type: "default" as const, rId: "rIdH" }],
    footerReferences: [{ type: "default" as const, rId: "rIdF" }],
  };
  return {
    package: {
      document: {
        content: [
          {
            type: "paragraph",
            content: [
              {
                type: "run",
                content: [
                  { type: "text", text: "body" },
                  { type: "footnoteRef", id: 1 },
                ],
              },
            ],
          },
        ],
        finalSectionProperties: page,
      },
      headers: new Map([
        [
          "rIdH",
          {
            type: "header",
            hdrFtrType: "default",
            content: storyContent(style),
          },
        ],
      ]),
      footers: new Map([
        [
          "rIdF",
          {
            type: "footer",
            hdrFtrType: "default",
            content: storyContent(style),
          },
        ],
      ]),
      footnotes: [{ type: "footnote", id: 1, content: storyContent(style) }],
      styles: {
        styles: [
          {
            styleId: "ListStyle",
            type: "paragraph",
            pPr: { numPr: { kind: "reference", numId: 1, ilvl: 0 } },
          },
        ],
      },
      numbering: {
        abstractNums: [
          {
            abstractNumId: 0,
            levels: [0, 1].map((ilvl) => ({
              ilvl,
              start: 1,
              numFmt:
                ilvl === 0 ? ("decimal" as const) : ("lowerLetter" as const),
              lvlText: `%${String(ilvl + 1)}.`,
              pPr: {
                indentLeft: 720 * (ilvl + 1),
                indentFirstLine: -360,
                hangingIndent: true,
              },
            })),
          },
        ],
        nums: [{ numId: 1, abstractNumId: 0 }],
      },
    },
  };
};

for (const style of [false, true])
  test(`canonical story projections read lists as the story conversion with numbering (${style ? "style" : "direct"})`, async () => {
    const bytes = await prepareCanonicalDocxInput(
      await createDocx(documentWith(style)),
    );
    const parsed = await parseDocx(bytes.unwrap(), {
      detectVariables: false,
      preloadFonts: false,
    });
    const session = createCanonicalSession(
      parsed,
      parsed.package.styles,
    ).unwrap();
    const document = readModel(session.document);
    const stories = [
      ...[...(document.package.headers ?? [])].map(([rId, story]) => ({
        story: { kind: "header" as const, rId },
        content: story.content,
      })),
      ...[...(document.package.footers ?? [])].map(([rId, story]) => ({
        story: { kind: "footer" as const, rId },
        content: story.content,
      })),
      ...(document.package.footnotes ?? [])
        .filter(
          ({ noteType }) => noteType === undefined || noteType === "normal",
        )
        .map((note) => ({
          story: { kind: "footnote" as const, id: note.id },
          content: note.content,
        })),
    ];
    expect(stories.length).toBe(3);
    const differences: unknown[] = [];
    for (const { story, content } of stories) {
      const projected = session.projectStory(story).unwrap().doc;
      const converted = headerFooterToProseDoc(structuredClone(content), {
        ...(document.package.styles != null && {
          styles: structuredClone(document.package.styles),
        }),
        ...(document.package.numbering != null && {
          numbering: structuredClone(document.package.numbering),
        }),
        ...(document.package.theme != null && {
          theme: structuredClone(document.package.theme),
        }),
      });
      expect(projected.childCount).toBe(converted.childCount);
      for (let index = 0; index < converted.childCount; index += 1) {
        const left: Record<string, unknown> = converted.child(index).attrs;
        const right: Record<string, unknown> = projected.child(index).attrs;
        for (const key of new Set([
          ...Object.keys(left),
          ...Object.keys(right),
        ]))
          if (JSON.stringify(left[key]) !== JSON.stringify(right[key]))
            differences.push({
              story,
              index,
              key,
              withNumbering: left[key],
              projected: right[key],
            });
      }
    }
    expect(differences).toEqual([]);
  });
