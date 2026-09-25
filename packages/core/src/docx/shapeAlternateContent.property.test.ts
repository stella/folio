/**
 * A shape read from `mc:AlternateContent` keeps its `mc:Fallback`.
 *
 * The text-box enrichment and the run parser model a shape from the Choice
 * branch, and the serializer rebuilt it as a bare `w:drawing`, so an unedited
 * save dropped the Fallback (a VML copy of the shape and its text) that
 * consumers without the Choice's namespace render instead (ECMA-376 Part 3).
 *
 * Two laws, over arbitrary text-box text and an optional edit, through the
 * model and through the editor projection:
 * - an unedited shape writes its element back whole;
 * - no save writes a Fallback whose text disagrees with its Choice.
 */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import JSZip from "jszip";
import { EditorState, TextSelection } from "prosemirror-state";

import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";

import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import { ExtensionManager } from "../prosemirror/extensions/ExtensionManager";
import { createStarterKit } from "../prosemirror/extensions/StarterKit";
import type { Document, Paragraph, ShapeContent } from "../types/document";
import { parseDocx } from "./parser";
import { createEmptyDocx, repackDocx } from "./rezip";

setDefaultTimeout(propertyTestTimeout(30_000));

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const NAMESPACES = {
  w: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
  wp: "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
  a: "http://schemas.openxmlformats.org/drawingml/2006/main",
  wps: "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
  w14: "http://schemas.microsoft.com/office/word/2010/wordml",
  mc: "http://schemas.openxmlformats.org/markup-compatibility/2006",
  v: "urn:schemas-microsoft-com:vml",
  o: "urn:schemas-microsoft-com:office:office",
} as const;
const ROOT_NAMESPACES = Object.entries(NAMESPACES)
  .map(([prefix, uri]) => ` xmlns:${prefix}="${uri}"`)
  .join("");

const anchor = (id: number, name: string, graphic: string): string =>
  '<w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" ' +
  'relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">' +
  '<wp:simplePos x="0" y="0"/>' +
  '<wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>' +
  '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>' +
  `<wp:extent cx="1828800" cy="457200"/><wp:effectExtent l="0" t="0" r="0" b="0"/>` +
  `<wp:wrapNone/><wp:docPr id="${id}" name="${name}"/>` +
  `<a:graphic><a:graphicData uri="${NAMESPACES.wps}">${graphic}</a:graphicData></a:graphic>` +
  "</wp:anchor></w:drawing>";

/**
 * Package paragraph ids, as hosts stamp them at ingest: the editor mints one
 * for a paragraph that has none, which is an edit to the text box.
 */
const HOST_PARA_ID = "2A000001";
const TEXT_BOX_PARA_ID = "2A000002";

const textBoxParagraph = (text: string): string =>
  `<w:p w14:paraId="${TEXT_BOX_PARA_ID}"><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;

const textBoxAlternateContent = (text: string): string =>
  '<mc:AlternateContent><mc:Choice Requires="wps">' +
  anchor(
    1,
    "Text Box 1",
    '<wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/>' +
      '<a:ext cx="1828800" cy="457200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>' +
      "</wps:spPr><wps:txbx><w:txbxContent>" +
      textBoxParagraph(text) +
      '</w:txbxContent></wps:txbx><wps:bodyPr rot="0" vert="horz"/></wps:wsp>',
  ) +
  '</mc:Choice><mc:Fallback><w:pict><v:shape id="Text Box 1" o:spid="_x0000_s1" ' +
  'type="#_x0000_t202" style="position:absolute;width:144pt;height:36pt">' +
  `<v:textbox><w:txbxContent>${textBoxParagraph(text)}` +
  "</w:txbxContent></v:textbox></v:shape></w:pict></mc:Fallback></mc:AlternateContent>";

const CONNECTOR_ALTERNATE_CONTENT =
  '<mc:AlternateContent><mc:Choice Requires="wps">' +
  anchor(
    2,
    "Straight Connector 2",
    '<wps:wsp><wps:cNvCnPr/><wps:spPr><a:xfrm><a:off x="0" y="0"/>' +
      '<a:ext cx="1828800" cy="0"/></a:xfrm><a:prstGeom prst="line"><a:avLst/></a:prstGeom>' +
      '<a:ln w="9525"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>' +
      "</wps:spPr><wps:bodyPr/></wps:wsp>",
  ) +
  '</mc:Choice><mc:Fallback><w:pict><v:line id="Straight Connector 2" o:spid="_x0000_s2" ' +
  'style="position:absolute" from="0,0" to="144pt,0" strokeweight=".75pt"/></w:pict>' +
  "</mc:Fallback></mc:AlternateContent>";

const documentXml = (runContent: string): string =>
  `${XML_DECLARATION}<w:document${ROOT_NAMESPACES}><w:body>` +
  `<w:p w14:paraId="${HOST_PARA_ID}"><w:r><w:t>Host</w:t></w:r><w:r>${runContent}</w:r></w:p>` +
  '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>';

const open = async (xml: string): Promise<Document> => {
  const zip = await JSZip.loadAsync(await createEmptyDocx());
  zip.file("word/document.xml", xml);
  return parseDocx(await zip.generateAsync({ type: "arraybuffer" }), { preloadFonts: false });
};

const savedDocumentPart = async (document: Document): Promise<string> => {
  const saved = await repackDocx(document, { updateModifiedDate: false });
  return (await (await JSZip.loadAsync(saved)).file("word/document.xml")?.async("text")) ?? "";
};

const throughEditor = (document: Document): Document =>
  fromProseDoc(toProseDoc(document), document);

const shapesOf = (document: Document): ShapeContent[] =>
  document.package.document.content
    .filter((block): block is Paragraph => block.type === "paragraph")
    .flatMap((paragraph) => paragraph.content)
    .flatMap((item) => (item.type === "run" ? item.content : []))
    .filter((content): content is ShapeContent => content.type === "shape");

const textOf = (xml: string): string =>
  [...xml.matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/gu)].map((match) => match[1]).join("");

const branchText = (alternateContent: string, branch: "Choice" | "Fallback"): string =>
  textOf(
    new RegExp(`<mc:${branch}\\b[\\s\\S]*?</mc:${branch}>`, "u").exec(alternateContent)?.[0] ?? "",
  );

const alternateContentsOf = (xml: string): string[] =>
  [...xml.matchAll(/<mc:AlternateContent\b[\s\S]*?<\/mc:AlternateContent>/gu)].map(
    (match) => match[0],
  );

/** A Fallback that says something else than its Choice misleads every consumer that reads it. */
const expectFallbacksAgree = (xml: string): void => {
  for (const alternateContent of alternateContentsOf(xml)) {
    expect(branchText(alternateContent, "Fallback")).toBe(branchText(alternateContent, "Choice"));
  }
};

/** Replace the text of the first text box's first run, the way an edit to the model would. */
const editTextBoxInModel = (document: Document, text: string): void => {
  const paragraph = shapesOf(document).at(0)?.shape.textBody?.content.at(0);
  const run = paragraph?.type === "paragraph" ? paragraph.content.at(0) : undefined;
  const content = run?.type === "run" ? run.content.at(0) : undefined;
  if (content?.type !== "text") {
    throw new Error("expected text in the text box");
  }
  content.text = text;
};

/** An editor state carrying the plugins the real editor runs. */
const editorStateOf = (document: Document): EditorState => {
  const manager = new ExtensionManager(createStarterKit());
  manager.buildSchema();
  manager.initializeRuntime();
  return EditorState.create({ doc: toProseDoc(document), plugins: manager.getPlugins() });
};

type EditTarget = "textBox" | "host";

/** Type at the start of the text box or of its host paragraph, the way a user would. */
const typeInEditor = (document: Document, target: EditTarget, text: string): Document => {
  const state = editorStateOf(document);
  let textPosition: number | undefined;
  state.doc.descendants((node, position) => {
    if (textPosition !== undefined) {
      return false;
    }
    if (node.type.name === (target === "textBox" ? "textBox" : "paragraph")) {
      textPosition = position + (target === "textBox" ? 2 : 1);
      return false;
    }
    return true;
  });
  if (textPosition === undefined) {
    throw new Error(`expected a ${target} node`);
  }
  const edited = state.apply(
    state.tr.setSelection(TextSelection.create(state.doc, textPosition)).insertText(text),
  );
  return fromProseDoc(edited.doc, document);
};

const safeText = fc.stringMatching(/^[A-Za-z0-9][A-Za-z0-9 ]{0,24}$/u);

describe("a shape read from mc:AlternateContent", () => {
  test("an unedited text box writes its element back whole, through the model and the editor", async () => {
    await fc.assert(
      fc.asyncProperty(safeText, async (text) => {
        const alternateContent = textBoxAlternateContent(text);
        const parsed = await open(documentXml(alternateContent));

        expect(await savedDocumentPart(parsed)).toContain(alternateContent);
        expect(await savedDocumentPart(throughEditor(parsed))).toContain(alternateContent);
      }),
      propertyConfig({ numRuns: 20 }),
    );
  }, 120_000);

  test("an edited text box is regenerated without the stale Fallback", async () => {
    await fc.assert(
      fc.asyncProperty(safeText, safeText, async (original, replacement) => {
        fc.pre(original !== replacement);
        const parsed = await open(documentXml(textBoxAlternateContent(original)));
        editTextBoxInModel(parsed, replacement);
        const saved = await savedDocumentPart(parsed);

        expect(saved).not.toContain("mc:Fallback");
        expect(textOf(saved)).toBe(`Host${replacement}`);
        expectFallbacksAgree(saved);
      }),
      propertyConfig({ numRuns: 20 }),
    );
  }, 120_000);

  test("typing into a text box in the editor drops the stale Fallback", async () => {
    await fc.assert(
      fc.asyncProperty(safeText, safeText, async (original, typed) => {
        const parsed = await open(documentXml(textBoxAlternateContent(original)));
        const saved = await savedDocumentPart(typeInEditor(parsed, "textBox", typed));

        expect(saved).not.toContain("mc:Fallback");
        expect(textOf(saved)).toBe(`Host${typed}${original}`);
        expectFallbacksAgree(saved);
      }),
      propertyConfig({ numRuns: 20 }),
    );
  }, 120_000);

  test("typing elsewhere in the document keeps the text box's element whole", async () => {
    await fc.assert(
      fc.asyncProperty(safeText, safeText, async (text, typed) => {
        const alternateContent = textBoxAlternateContent(text);
        const parsed = await open(documentXml(alternateContent));
        const saved = await savedDocumentPart(typeInEditor(parsed, "host", typed));

        expect(saved).toContain(alternateContent);
        expect(textOf(saved)).toBe(`${typed}Host${text}${text}`);
      }),
      propertyConfig({ numRuns: 20 }),
    );
  }, 120_000);

  test("a shape without a text body keeps its Fallback until it is edited", async () => {
    const parsed = await open(documentXml(CONNECTOR_ALTERNATE_CONTENT));
    expect(await savedDocumentPart(parsed)).toContain(CONNECTOR_ALTERNATE_CONTENT);
    expect(await savedDocumentPart(throughEditor(parsed))).toContain(CONNECTOR_ALTERNATE_CONTENT);

    const connector = shapesOf(parsed).at(0);
    if (!connector) {
      throw new Error("expected the connector shape");
    }
    connector.shape.size = { ...connector.shape.size, width: connector.shape.size.width * 2 };
    const saved = await savedDocumentPart(parsed);
    expect(saved).not.toContain("mc:AlternateContent");
    expect(saved).toContain('<wp:extent cx="3657600"');
  });

  test("a branch holding more than the shape is not replayed in its place", async () => {
    // Two text boxes in one Choice become two shapes; replaying the element
    // for each would write both boxes twice.
    const choice = textBoxAlternateContent("first").replace(
      "</mc:Choice>",
      `${/<w:drawing>[\s\S]*<\/w:drawing>/u.exec(textBoxAlternateContent("second"))?.[0] ?? ""}</mc:Choice>`,
    );
    const parsed = await open(documentXml(choice));

    expect(shapesOf(parsed).map((content) => content.alternateContent)).toEqual([
      undefined,
      undefined,
    ]);
    const saved = await savedDocumentPart(parsed);
    expect(textOf(saved)).toBe("Hostfirstsecond");
  });
});
