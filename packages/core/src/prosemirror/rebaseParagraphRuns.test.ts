import { expect, test } from "bun:test";
import fc from "fast-check";
import { EditorState, TextSelection } from "prosemirror-state";

import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";
import { createDocx } from "../docx/rezip";
import { parseDocx } from "../docx/parser";
import type { Document } from "../types/document";
import { fromProseDoc } from "./conversion/fromProseDoc";
import { toProseDoc } from "./conversion/toProseDoc";
import { createDocumentStylesPlugin } from "./plugins/documentStyles";
import { singletonManager, schema } from "./schema";
import { rebaseParagraphRunContent, rebaseParagraphRuns } from "./rebaseParagraphRuns";
import { expectFontSizeMarkAttrs, expectTextColorMarkAttrs } from "./attrs";

type LinkedDocumentOptions = { color: string | undefined; paragraphStyleId: "Normal" | "Heading1" };
const linkedDocument = ({ color, paragraphStyleId }: LinkedDocumentOptions): Document => ({
  package: {
    styles: {
      docDefaults: { rPr: { fontSize: 22 } },
      styles: [
        { styleId: "Normal", type: "paragraph", default: true },
        {
          styleId: "Heading1",
          type: "paragraph",
          basedOn: "Normal",
          rPr: { bold: true, fontSize: 32 },
        },
        {
          styleId: "LinkBase",
          type: "character",
          rPr: { color: { rgb: "0563C1" }, underline: { style: "single" } },
        },
        { styleId: "LinkLook", type: "character", basedOn: "LinkBase", rPr: { italic: true } },
      ],
    },
    document: {
      content: [
        {
          type: "paragraph",
          formatting: { styleId: paragraphStyleId },
          content: [
            {
              type: "hyperlink",
              href: "https://example.test",
              children: [
                {
                  type: "run",
                  formatting: {
                    styleId: "LinkLook",
                    bold: false,
                    ...(color ? { color: { rgb: color } } : {}),
                  },
                  content: [{ type: "text", text: "linked text" }],
                },
              ],
            },
          ],
        },
      ],
    },
  },
});

test(
  "style commands preserve hyperlink character cascades and authored overrides across save",
  async () => {
    await assertProperty(
      fc.asyncProperty(
        fc.constantFrom("applyStyle", "clearStyle", "clearTextColor"),
        fc.option(fc.constantFrom("AA0000", "00AA00"), { nil: undefined }),
        async (operation, directColor) => {
          const base = await parseDocx(
            await createDocx(
              linkedDocument({
                color: directColor,
                paragraphStyleId: operation === "clearStyle" ? "Heading1" : "Normal",
              }),
            ),
            {
              preloadFonts: false,
              detectVariables: false,
            },
          );
          const doc = toProseDoc(base);
          let state = EditorState.create({
            doc,
            selection: TextSelection.create(doc, 1, doc.content.size - 1),
            plugins: [createDocumentStylesPlugin(base.package.styles)],
          });
          const applyStyle = singletonManager.getCommand("applyStyle");
          const clearStyle = singletonManager.getCommand("clearStyle");
          const clearColor = singletonManager.getCommand("clearTextColor");
          if (!applyStyle || !clearStyle || !clearColor) throw new Error("Missing style commands");
          const command = {
            applyStyle: applyStyle("Heading1"),
            clearStyle: clearStyle(),
            clearTextColor: clearColor(),
          }[operation];
          expect(
            command(state, (tr) => {
              state = state.apply(tr);
            }),
          ).toBe(true);
          const check = (current: typeof doc) => {
            const node = current.firstChild?.firstChild;
            expect(
              node?.marks.find(({ type }) => type.name === "characterStyle")?.attrs["styleId"],
            ).toBe("LinkLook");
            expect(node?.marks.some(({ type }) => type.name === "hyperlink")).toBe(true);
            expect(node?.marks.some(({ type }) => type.name === "underline")).toBe(true);
            expect(node?.marks.some(({ type }) => type.name === "italic")).toBe(true);
            expect(node?.marks.some(({ type }) => type.name === "bold")).toBe(false);
            const color = node?.marks.find(({ type }) => type.name === "textColor");
            expect(color ? expectTextColorMarkAttrs(color).rgb : undefined).toBe(
              operation === "clearTextColor" ? "0563C1" : (directColor ?? "0563C1"),
            );
          };
          check(state.doc);
          const reopened = await parseDocx(await createDocx(fromProseDoc(state.doc, base)), {
            preloadFonts: false,
            detectVariables: false,
          });
          check(toProseDoc(reopened));
        },
      ),
      { numRuns: 24 },
    );
  },
  propertyTestTimeout(30_000),
);

test(
  "an unchanged cascade preserves run carriers and their authored provenance",
  () => {
    assertProperty(
      fc.property(fc.boolean(), fc.boolean(), fc.boolean(), (bold, rtl, explicit) => {
        const marks = [
          ...(bold ? [schema.mark("bold")] : []),
          ...(rtl ? [schema.mark("rtl")] : []),
          ...(explicit
            ? [schema.mark("runFormattingOverride", { _authoredOn: ["bold", "rtl"] })]
            : []),
        ];
        const paragraph = schema.nodes["paragraph"]!.create({ paraId: "source" }, [
          schema.text("Text", marks),
          schema.nodes["tab"]!.create(null, null, marks),
        ]);
        const target = schema.nodes["paragraph"]!.create({ paraId: "target" });
        let rebases = 0;
        const content = rebaseParagraphRunContent({
          paragraph,
          target,
          position: 0,
          styleResolver: null,
          onRebased: () => {
            rebases += 1;
          },
        });
        expect(content).toBe(paragraph.content);
        expect(rebases).toBe(0);
        const state = EditorState.create({ doc: schema.node("doc", null, paragraph) });
        const tr = rebaseParagraphRuns({
          tr: state.tr,
          position: 0,
          previous: paragraph,
          target,
          styleResolver: null,
        });
        expect(tr.steps).toHaveLength(0);
        expect(tr.doc).toBe(state.doc);
      }),
      { numRuns: 32 },
    );
  },
  propertyTestTimeout(30_000),
);

test(
  "paragraph style commands carry pending cursor formatting through changed and equal cascades",
  async () => {
    await assertProperty(
      fc.asyncProperty(
        fc.constantFrom("applyStyle", "clearStyle", "sameStyle"),
        fc.constantFrom(1, 2, 4),
        async (operation, cursor) => {
          const model = linkedDocument({ color: undefined, paragraphStyleId: "Normal" });
          model.package.document.content = [
            {
              type: "paragraph",
              formatting: { styleId: operation === "clearStyle" ? "Heading1" : "Normal" },
              content: [{ type: "run", content: [{ type: "text", text: "abc" }] }],
            },
          ];
          const base = await parseDocx(await createDocx(model), {
            preloadFonts: false,
            detectVariables: false,
          });
          const doc = toProseDoc(base);
          let state = EditorState.create({
            doc,
            selection: TextSelection.create(doc, cursor),
            plugins: [createDocumentStylesPlugin(base.package.styles)],
          });
          const italic = singletonManager.requireCommand("toggleItalic")();
          expect(
            italic(state, (tr) => {
              state = state.apply(tr);
            }),
          ).toBe(true);
          expect(state.storedMarks?.some((mark) => mark.type.name === "italic")).toBe(true);
          const command =
            operation === "clearStyle"
              ? singletonManager.requireCommand("clearStyle")()
              : singletonManager.requireCommand("applyStyle")(
                  operation === "sameStyle" ? "Normal" : "Heading1",
                );
          expect(
            command(state, (tr) => {
              state = state.apply(tr);
            }),
          ).toBe(true);
          const marks = state.storedMarks ?? [];
          expect(marks.some((mark) => mark.type.name === "italic")).toBe(true);
          expect(marks.some((mark) => mark.type.name === "bold")).toBe(operation === "applyStyle");
          const size = marks.find((mark) => mark.type.name === "fontSize");
          if (!size) throw new Error("Stored font size is missing");
          expect(expectFontSizeMarkAttrs(size).size).toBe(operation === "applyStyle" ? 32 : 22);
        },
      ),
      { numRuns: 27 },
    );
  },
  propertyTestTimeout(30_000),
);
