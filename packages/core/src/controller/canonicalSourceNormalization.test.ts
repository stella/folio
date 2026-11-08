import { describe, expect, test } from "bun:test";
import { EditorState, TextSelection } from "prosemirror-state";

import type { Document } from "@stll/docx-core";
import { ZipArchive } from "@stll/docx-core/zip";

import { serializeCanonicalSave } from "../docx/canonicalSave";
import { prepareCanonicalDocxInput } from "../docx/canonicalSessionInput";
import { parseDocx } from "../docx/parser";
import { RELATIONSHIP_TYPES } from "../docx/relsParser";
import { schema } from "../prosemirror/schema";
import { createCanonicalSession } from "./canonicalSession";
import {
  canonicalSourceRestores,
  normalizeCanonicalSource,
  withUneditedSourceRestored,
} from "./canonicalSourceNormalization";

const XML_DECLARATION =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const RELATIONSHIPS =
  "http://schemas.openxmlformats.org/package/2006/relationships";
const NAMESPACES =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"';
const PARTS = [
  "word/document.xml",
  "word/header1.xml",
  "word/footer1.xml",
] as const;

/**
 * A body paragraph and a header paragraph whose text holds a literal tab, and
 * a footer with no blocks at all.
 */
const sourceBytes = async (): Promise<Uint8Array> => {
  const zip = new ZipArchive();
  zip.set(
    "[Content_Types].xml",
    `${XML_DECLARATION}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/><Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/></Types>`,
  );
  zip.set(
    "_rels/.rels",
    `${XML_DECLARATION}<Relationships xmlns="${RELATIONSHIPS}"><Relationship Id="rId1" Type="${RELATIONSHIP_TYPES.officeDocument}" Target="word/document.xml"/></Relationships>`,
  );
  zip.set(
    "word/_rels/document.xml.rels",
    `${XML_DECLARATION}<Relationships xmlns="${RELATIONSHIPS}"><Relationship Id="rId20" Type="${RELATIONSHIP_TYPES.header}" Target="header1.xml"/><Relationship Id="rId21" Type="${RELATIONSHIP_TYPES.footer}" Target="footer1.xml"/></Relationships>`,
  );
  zip.set(
    "word/document.xml",
    `${XML_DECLARATION}<w:document ${NAMESPACES}><w:body><w:p w14:paraId="10000001"><w:r><w:t xml:space="preserve">Name\tValue</w:t></w:r></w:p><w:p w14:paraId="10000002"><w:r><w:t>Body</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rId20"/><w:footerReference w:type="default" r:id="rId21"/></w:sectPr></w:body></w:document>`,
  );
  zip.set(
    "word/header1.xml",
    `${XML_DECLARATION}<w:hdr ${NAMESPACES}><w:p w14:paraId="20000001"><w:r><w:t xml:space="preserve">Left\tRight</w:t></w:r></w:p></w:hdr>`,
  );
  zip.set("word/footer1.xml", `${XML_DECLARATION}<w:ftr ${NAMESPACES}/>`);
  return zip.generate({ compression: "store" });
};

const partsOf = async (bytes: ArrayBuffer | Uint8Array) => {
  const zip = await ZipArchive.read(bytes);
  return Object.fromEntries(
    await Promise.all(
      PARTS.map(async (part) => [part, await zip.entry(part)?.text()] as const),
    ),
  );
};

const parsed = async (bytes: Uint8Array): Promise<Document> =>
  parseDocx((await prepareCanonicalDocxInput(bytes)).unwrap(), {
    preloadFonts: false,
    detectVariables: false,
  });

const saved = async (session: ReturnType<typeof createCanonicalSession>) =>
  (
    await serializeCanonicalSave({
      snapshot: session.unwrap().captureSaveSnapshot(),
      featureFlags: { selectiveSave: true },
    })
  ).buffer;

describe("source spellings the canonical session reads in its own form", () => {
  test("a literal tab and an empty footer open, and an unedited save gives every part back as read", async () => {
    const bytes = await sourceBytes();
    const session = createCanonicalSession(await parsed(bytes));
    expect(session.isOk()).toBe(true);
    expect(await partsOf(await saved(session))).toEqual(await partsOf(bytes));
  });

  test("an edit in the tab's paragraph saves the tab as a tab element, the other parts as read", async () => {
    const bytes = await sourceBytes();
    const session = createCanonicalSession(await parsed(bytes)).unwrap();
    let state = EditorState.create({ schema, doc: session.projection.doc });
    state = state.apply(
      state.tr.setSelection(TextSelection.create(state.doc, 1)),
    );
    const commit = session
      .prepareReplace(state, { from: 1, to: 1, text: "X" })
      .unwrap();
    expect(commit.publish().isOk()).toBe(true);
    const source = await partsOf(bytes);
    const result = await partsOf(
      (
        await serializeCanonicalSave({
          snapshot: session.captureSaveSnapshot(),
          featureFlags: { selectiveSave: true },
        })
      ).buffer,
    );
    // The edited paragraph's runs, as the session writes them.
    const runs = /<w:p [^>]*w14:paraId="10000001"[^>]*>(.*?)<\/w:p>/su.exec(
      result["word/document.xml"] ?? "",
    )?.[1];
    expect(runs?.replaceAll(/<w:rPr>.*?<\/w:rPr>/gu, "")).toBe(
      "<w:r><w:t>X</w:t></w:r><w:r><w:t>Name</w:t><w:tab/><w:t>Value</w:t></w:r>",
    );
    expect(result["word/header1.xml"]).toBe(source["word/header1.xml"]);
    expect(result["word/footer1.xml"]).toBe(source["word/footer1.xml"]);
  });

  test("a story edited since the session opened keeps the session's spelling", async () => {
    const document = await parsed(await sourceBytes());
    const source = normalizeCanonicalSource(document);
    const restores = canonicalSourceRestores(
      source.document,
      source.normalized,
    );
    expect(restores.map(({ story }) => story)).toEqual([
      "main",
      { kind: "header", rId: "rId20" },
      { kind: "footer", rId: "rId21" },
    ]);
    const footer = source.document.package.footers?.get("rId21");
    const [paragraph] = footer?.content ?? [];
    if (footer === undefined || paragraph?.type !== "paragraph")
      throw new Error("The empty footer has no paragraph.");
    const typed = new Map(source.document.package.footers);
    typed.set("rId21", {
      ...footer,
      content: [
        {
          ...paragraph,
          content: [
            { type: "run", content: [{ type: "text", text: "Typed" }] },
          ],
        },
      ],
    });
    const edited: Document = {
      ...source.document,
      package: { ...source.document.package, footers: typed },
    };
    const restored = withUneditedSourceRestored(edited, restores);
    // The edited footer keeps its paragraph; the unedited header reads as before.
    expect(restored.package.footers?.get("rId21")?.content).toHaveLength(1);
    expect(restored.package.headers?.get("rId20")?.content).toEqual(
      document.package.headers?.get("rId20")?.content,
    );
    expect(
      withUneditedSourceRestored(
        source.document,
        restores,
      ).package.footers?.get("rId21")?.content,
    ).toEqual([]);
  });
});
