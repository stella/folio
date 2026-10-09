import { BODY_TEXT_OUTLINE_LEVEL } from "@stll/docx-core/model";

import type { Style } from "../../types/document";
import { paragraphNumberingAttr } from "../../prosemirror/numberingAttr";
import { schema } from "../../prosemirror/schema";

export const HEADING_COLLECTOR_STYLES: Style[] = [
  { type: "paragraph", styleId: "Normal", name: "Normal" },
  { type: "paragraph", styleId: "Heading1", name: "heading 1" },
  { type: "paragraph", styleId: "Heading2", name: "heading 2" },
  { type: "paragraph", styleId: "CustomHeading", name: "Clause heading", basedOn: "Heading2" },
  {
    type: "paragraph",
    styleId: "BodyClause",
    name: "Body clause",
    basedOn: "Heading2",
    pPr: { outlineLevel: BODY_TEXT_OUTLINE_LEVEL },
  },
  { type: "paragraph", styleId: "ExplicitOutline", name: "Custom outline" },
  { type: "paragraph", styleId: "ListParagraph", name: "List Paragraph" },
  { type: "paragraph", styleId: "BodyText", name: "Body Text" },
];

const paragraph = (styleId: string, text: string, attrs: Record<string, unknown> = {}) =>
  schema.node("paragraph", { styleId, ...attrs }, [schema.text(text)]);

const boldParagraph = (styleId: string, text: string, attrs: Record<string, unknown> = {}) =>
  schema.node("paragraph", { styleId, ...attrs }, [schema.text(text, [schema.mark("bold")])]);

const definedTermParagraph = () =>
  schema.node("paragraph", { styleId: "Normal" }, [
    schema.text('"Synthetic Term"', [schema.mark("bold")]),
    schema.text(" means a term used only by this synthetic document fixture."),
  ]);

const generatedHeadings = Array.from({ length: 64 }, (_, index) => {
  const level = index % 2 === 0 ? 1 : 2;
  return paragraph(
    `Heading${level}`,
    `Section ${index + 1}: General provisions governing synthetic agreements and related obligations`,
  );
});

export const HEADING_COLLECTOR_DOCUMENT = schema.node("doc", null, [
  ...generatedHeadings,
  paragraph("CustomHeading", "Inherited clause heading"),
  paragraph("BodyClause", "Inherited heading overridden to body text"),
  paragraph("ExplicitOutline", "Explicit outline", {
    outlineLevel: { kind: "heading", level: 2 },
  }),
  ...[0, 1, 2].map((ilvl) =>
    paragraph(`ListParagraph`, `Numbered list item at level ${ilvl}`, {
      numPr: paragraphNumberingAttr({ kind: "reference", numId: 1, ilvl }),
      listMarker: `${ilvl + 1}.`,
    }),
  ),
  definedTermParagraph(),
  boldParagraph("Normal", "PLAIN BOLD CAPS"),
  paragraph("Normal", "The parties agree that this ordinary clause remains body text."),
  paragraph("Heading1", "Body override", {
    outlineLevel: BODY_TEXT_OUTLINE_LEVEL,
  }),
]);

export const HEADING_COLLECTOR_RUN_IN_DOCUMENT = schema.node("doc", null, [
  schema.node("paragraph", { styleId: "Heading1" }, [
    schema.text("Run-in prefix: "),
    schema.text("Article 7. ", [schema.mark("bold"), schema.mark("runIdentity", { id: 1 })]),
    schema.text("Scope", [schema.mark("bold"), schema.mark("runIdentity", { id: 2 })]),
    schema.text(" The parties agree to the following terms."),
    schema.text(" Later emphasized words", [schema.mark("bold")]),
  ]),
  paragraph("Heading2", "Single run remains the fallback title"),
  schema.node("paragraph", { styleId: "Heading1" }),
  schema.node("paragraph", { styleId: "Heading1" }, [
    schema.text("   ", [schema.mark("bold")]),
    schema.text("Visible fallback after an empty formatted span"),
  ]),
  schema.node("paragraph", { styleId: "Heading2" }, [
    schema.text("Plain prefix: "),
    schema.text("Underlined title", [
      schema.mark("underline"),
      schema.mark("runIdentity", { id: 3 }),
    ]),
    schema.text(" and more", [schema.mark("underline"), schema.mark("runIdentity", { id: 4 })]),
    schema.text(" body follows"),
  ]),
  schema.node("paragraph", { styleId: "Heading1" }, [
    schema.text("Bold run", [schema.mark("bold")]),
    schema.text(" Underline run", [schema.mark("underline")]),
    schema.text(" unformatted body"),
  ]),
  schema.node("paragraph", { styleId: "Heading2" }, [
    schema.text("Italic run-in title", [
      schema.mark("italic"),
      schema.mark("runIdentity", { id: 5 }),
    ]),
    schema.text(" and body text"),
  ]),
]);
