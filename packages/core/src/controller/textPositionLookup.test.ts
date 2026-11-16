import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import { OP_STORIES, type TextPosition } from "@stll/docx-core/ops";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import { createCanonicalSession } from "../canonical/canonicalSession";
import { fromMarkdown } from "../markdown/fromMarkdown";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import { schema } from "../prosemirror/schema";
import { findParagraphByParaId } from "../prosemirror/utils/findParagraphByParaId";
import type { Document, ParagraphContent } from "../types/document";
import { readModel } from "../types/documentModel";
import { legacyTextAddressAt } from "./legacyTextAddress";
import { legacyEditorPositionAt } from "./textPositionLookup";

/** A document from Markdown whose paragraphs carry ids, as a parsed DOCX's do. */
const withParagraphIds = (markdown: string) => {
  const document = fromMarkdown(markdown);
  for (const [index, block] of document.package.document.content.entries())
    if (block.type === "paragraph")
      block.paraId = (index + 1).toString(16).toUpperCase().padStart(8, "0");
  return document;
};

const openSession = (markdown: string) => {
  const created = createCanonicalSession(withParagraphIds(markdown));
  if (created.isErr()) throw created.error;
  return created.value;
};

const bodyParagraphIds = (session: ReturnType<typeof openSession>) =>
  readModel(session.document).package.document.content.flatMap((block) =>
    block.type === "paragraph" && block.paraId !== undefined
      ? [{ id: block.paraId }]
      : [],
  );

describe("legacyEditorPositionAt", () => {
  test("names the same text the canonical projection names at every offset", () => {
    const session = openSession(
      "Alpha **bold** and *slanted* text\n\nSecond paragraph here",
    );
    const legacy = toProseDoc(readModel(session.document));
    const projection = session.projection;
    const ids = bodyParagraphIds(session);
    expect(ids.length).toBe(2);
    for (const { id } of ids) {
      const paragraphText =
        findParagraphByParaId(legacy, id)?.node.textContent ?? "";
      expect(paragraphText.length).toBeGreaterThan(0);
      for (let offset = 0; offset <= paragraphText.length; offset += 1) {
        const position: TextPosition = {
          story: OP_STORIES.MAIN,
          blockId: id,
          offset,
        };
        const legacyAt = legacyEditorPositionAt(legacy, position);
        const canonicalAt = projection.positionAt(position);
        expect(canonicalAt.isOk()).toBe(true);
        if (legacyAt === null || canonicalAt.isErr())
          throw new Error(`No position for ${id}@${offset}`);
        const before = (doc: typeof legacy, at: number) =>
          doc.resolve(at).parent.textBetween(0, doc.resolve(at).parentOffset);
        expect(before(legacy, legacyAt)).toBe(paragraphText.slice(0, offset));
        expect(before(projection.doc, canonicalAt.value)).toBe(
          paragraphText.slice(0, offset),
        );
      }
    }
  });

  test("refuses an unknown paragraph, an offset past the end and another story", () => {
    const session = openSession("Only paragraph");
    const legacy = toProseDoc(readModel(session.document));
    const [first] = bodyParagraphIds(session);
    if (first === undefined) throw new Error("No paragraph id");
    const at = (position: TextPosition) =>
      legacyEditorPositionAt(legacy, position);
    expect(
      at({ story: OP_STORIES.MAIN, blockId: "NOSUCHID", offset: 0 }),
    ).toBeNull();
    expect(
      at({ story: OP_STORIES.MAIN, blockId: first.id, offset: 99 }),
    ).toBeNull();
    expect(
      at({ story: OP_STORIES.MAIN, blockId: first.id, offset: -1 }),
    ).toBeNull();
    expect(
      at({
        story: { kind: "footnote", id: 1 },
        blockId: first.id,
        offset: 0,
      }),
    ).toBeNull();
  });

  test("a bookmark boundary takes no logical unit", () => {
    const doc = schema.node("doc", null, [
      schema.node("paragraph", { paraId: "0000000A" }, [
        schema.node("bookmarkBoundary", { type: "start", id: 1 }),
        schema.text("A"),
      ]),
    ]);
    // The paragraph opens at 0: the boundary sits at 1, "A" at 2.
    expect(
      legacyEditorPositionAt(doc, {
        story: OP_STORIES.MAIN,
        blockId: "0000000A",
        offset: 1,
      }),
    ).toBe(3);
  });
});

setDefaultTimeout(propertyTestTimeout(30_000));

test("legacy text addresses round-trip actual source marker gaps and inline carriers", () => {
  const text = fc
    .array(fc.constantFrom("a", "b", "c"), { minLength: 1, maxLength: 5 })
    .map((chars) => chars.join(""));
  assertProperty(
    fc.property(text, text, (left, right) => {
      for (const nested of [false, true])
        for (const tab of [false, true]) {
          const content: ParagraphContent[] = [
            {
              type: "run",
              content: [
                ...(tab ? [{ type: "tab" } as const] : []),
                { type: "text", text: left },
              ],
            },
            { type: "bookmarkStart", id: 1, name: "_Ref" },
            { type: "bookmarkEnd", id: 1 },
            { type: "run", content: [{ type: "text", text: right }] },
          ];
          const source = {
            package: {
              document: {
                content: [
                  {
                    type: "paragraph",
                    paraId: "ABCDEF12",
                    content: nested
                      ? [
                          {
                            type: "inlineSdt",
                            properties: { sdtType: "richText", id: 7 },
                            content,
                          },
                        ]
                      : content,
                  },
                ],
              },
            },
          } satisfies Document;
          const doc = toProseDoc(source);
          const seam = left.length + Number(tab);
          for (let offset = 0; offset <= seam + right.length; offset += 1) {
            for (const zeroWidthBefore of offset === seam
              ? [0, 1, 2]
              : [undefined]) {
              const at = {
                story: OP_STORIES.MAIN,
                blockId: "ABCDEF12",
                offset,
                ...(zeroWidthBefore === undefined ? {} : { zeroWidthBefore }),
              };
              const position = legacyEditorPositionAt(doc, at);
              expect(position).not.toBeNull();
              if (position !== null)
                expect(legacyTextAddressAt(doc, position)).toEqual(at);
            }
          }
          const afterMarkers = {
            story: OP_STORIES.MAIN,
            blockId: "ABCDEF12",
            offset: seam,
          };
          expect(legacyEditorPositionAt(doc, afterMarkers)).toBe(
            legacyEditorPositionAt(doc, {
              ...afterMarkers,
              zeroWidthBefore: 2,
            }),
          );
          expect(
            legacyEditorPositionAt(doc, {
              ...afterMarkers,
              zeroWidthBefore: 3,
            }),
          ).toBeNull();
        }
    }),
    { numRuns: 20, id: "legacy-text-address-marker-gaps" },
  );
});
