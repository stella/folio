import { describe, expect, spyOn, test } from "bun:test";
import fc from "fast-check";
import { requiresXmlSpacePreserve } from "@stll/docx-core";

import { assertProperty } from "../../../../test/property-testing";
import { DOCX_CONFORMANCE_CLASSES } from "@stll/docx-core/model";
import {
  getSourceReplayToken,
  inheritSourceReplayToken,
  paragraphLogicalText,
} from "@stll/docx-core/ops";
import type { Document } from "../types/document";
import { parseDocumentBody, parseDocumentBodyTree } from "./documentParser";
import { standalonePreviewLedger } from "./previewBudget";
import { serializeDocument } from "./serializer/documentSerializer";
import { replayDocumentSource, trackDocumentSource } from "./documentSource";
import { assignDocumentParagraphPropertySourceContract } from "./paragraphPropertySource";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { EditorState } from "prosemirror-state";
import * as verbatimCapture from "./verbatimCapture";
import { getNamespaceUri, parseXml, parseXmlDocument } from "./xmlParser";
import { parseStreamingXmlWithSourceRanges } from "./streamingXmlParser";

const trackedSourceTree = (xml: string) => {
  const parsed = parseStreamingXmlWithSourceRanges(xml);
  return parsed.status === "parsed" ? parsed.value : parseXml(xml);
};

const PROFILES = [
  {
    conformance: DOCX_CONFORMANCE_CLASSES.TRANSITIONAL,
    uri: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
  },
  {
    conformance: DOCX_CONFORMANCE_CLASSES.STRICT,
    uri: "http://purl.oclc.org/ooxml/wordprocessingml/main",
  },
] as const;

const documentFor = (
  xml: string,
  conformance: Document["package"]["conformanceClass"],
  sourceReplay: "tracked" | "untracked" = "tracked",
): Document => {
  const document: Document = {
    package: {
      document: parseDocumentBodyTree({
        xml,
        doc: sourceReplay === "tracked" ? trackedSourceTree(xml) : parseXml(xml),
        styles: null,
        theme: null,
        numbering: null,
        rels: null,
        media: null,
        previews: standalonePreviewLedger(),
        sourceReplay,
      }),
      conformanceClass: conformance,
    },
  };
  assignDocumentParagraphPropertySourceContract(document, "a".repeat(64));
  if (sourceReplay === "tracked") trackDocumentSource(document);
  return document;
};

const serializeTracked = (document: Document) => {
  const sourceReplay = getSourceReplayToken(document);
  return serializeDocument(document, sourceReplay === undefined ? {} : { sourceReplay });
};

/** The old oracle compared only modelled values, so root markup loss was invisible. */
describe("Document source ownership", () => {
  test.each(["<other/>", "text", "<unbound:other/>"])(
    "tolerant parsing cannot authorize source replay with an extra root: %s",
    (extra) => {
      const xml = `<w:document xmlns:w="${PROFILES[0].uri}"><w:body><w:p/></w:body></w:document>${extra}`;
      expect(
        documentFor(xml, DOCX_CONFORMANCE_CLASSES.TRANSITIONAL).package.document.source,
      ).toBeUndefined();
    },
  );

  test("no edit retains every source character across profiles and namespace aliases", async () => {
    await assertProperty(
      fc.property(
        fc.constantFrom(...PROFILES),
        fc.constantFrom("w", "x", ""),
        fc.constantFrom(" ", "\n", "\r\n"),
        (profile, prefix, gap) => {
          const name = (local: string) => (prefix ? `${prefix}:${local}` : local);
          const binding = prefix ? `xmlns:${prefix}` : "xmlns";
          const xml = `<?xml version="1.0"?>${gap}<${name("document")} ${binding}="${profile.uri}" xmlns:e="urn:extension" e:flag="keep"><e:before/><${name("body")} e:body="keep">${gap}<${name("p")}><${name("r")}><${name("t")}>Text</${name("t")}></${name("r")}></${name("p")}>${gap}<e:opaque e:val="1 &gt; 0"/></${name("body")}><e:after/></${name("document")}>`;
          expect(serializeTracked(documentFor(xml, profile.conformance))).toBe(xml);
          // Import repair now hands the parser a retained tree. Exercise that
          // boundary too: losing its source XML must never disable capture.
          const retained: Document = {
            package: {
              conformanceClass: profile.conformance,
              document: parseDocumentBodyTree({
                sourceReplay: "tracked",
                xml,
                doc: trackedSourceTree(xml),
                styles: null,
                theme: null,
                numbering: null,
                rels: null,
                media: null,
                previews: standalonePreviewLedger(),
              }),
            },
          };
          trackDocumentSource(retained);
          expect(retained.package.document.source?.xml).toBe(xml);
          expect(serializeTracked(retained)).toBe(xml);
        },
      ),
      {},
    );
  });

  test.each(PROFILES)("one edit retains surrounding markup in $conformance", (profile) => {
    const untouched = '<x:p custom="keep"><x:r><x:t>Same</x:t></x:r></x:p>';
    const touched = "<x:p><x:r><x:t>Change</x:t></x:r></x:p>";
    const before = `<x:document xmlns:x="${profile.uri}" xmlns:e="urn:extension" e:flag="keep"><e:before/><x:body custom="keep">\n${untouched}\n`;
    const after =
      '\n<e:opaque/><x:sectPr><x:pgSz x:w="11906" x:h="16838"/></x:sectPr></x:body><e:after/></x:document>';
    const original = documentFor(before + touched + after, profile.conformance);
    const projected = toProseDoc(original);
    const state = EditorState.create({ doc: projected });
    const document = fromProseDoc(
      state.tr.insertText(" edited", projected.child(0).nodeSize + 1).doc,
      original,
    );
    const saved = serializeTracked(document);
    expect(saved.startsWith(before)).toBe(true);
    expect(saved.endsWith(after)).toBe(true);
    expect(saved).toContain(" edited");
    const root = parseXmlDocument(saved.slice(before.length, saved.length - after.length));
    expect(getNamespaceUri(root!)).toBe(profile.uri);
  });

  test("deleting or moving source records never replays the old sequence", () => {
    const xml = `<w:document xmlns:w="${PROFILES[0].uri}"><w:body><w:p><w:r><w:t>First</w:t></w:r></w:p><w:p><w:r><w:t>Second</w:t></w:r></w:p></w:body></w:document>`;
    const document = documentFor(xml, DOCX_CONFORMANCE_CLASSES.TRANSITIONAL, "untracked");
    document.package.document.content.shift();
    expect(serializeTracked(document)).not.toContain("First");
    const moved = documentFor(xml, DOCX_CONFORMANCE_CLASSES.TRANSITIONAL, "untracked");
    moved.package.document.content.reverse();
    const saved = serializeTracked(moved);
    expect(saved.indexOf("Second")).toBeLessThan(saved.indexOf("First"));
  });

  test("tracked immutable deletions and moves cannot replay the captured sequence", () => {
    const xml = `<w:document xmlns:w="${PROFILES[0].uri}"><w:body><w:p><w:r><w:t>First</w:t></w:r></w:p><w:p><w:r><w:t>Second</w:t></w:r></w:p></w:body></w:document>`;
    const original = documentFor(xml, DOCX_CONFORMANCE_CLASSES.TRANSITIONAL);
    for (const content of [
      original.package.document.content.slice(1),
      original.package.document.content.toReversed(),
    ]) {
      const changed = {
        ...original,
        package: { ...original.package, document: { ...original.package.document, content } },
      };
      inheritSourceReplayToken(changed, original);
      expect(
        replayDocumentSource({
          document: changed,
          token: getSourceReplayToken(changed),
          serialize: () => "",
        }),
      ).toBeNull();
      const saved = serializeTracked(changed);
      if (content.length === 1) expect(saved).not.toContain("First");
      else expect(saved.indexOf("Second")).toBeLessThan(saved.indexOf("First"));
    }
  });

  test("source replay performs required whitespace repairs across profiles and clones", async () => {
    await assertProperty(
      fc.property(
        fc.constantFrom(...PROFILES),
        fc.constantFrom("w", "x", ""),
        fc.constantFrom(" ", "\t", "\n"),
        fc.constantFrom("absent", "default", "preserve", "foreign"),
        fc.boolean(),
        (profile, prefix, whitespace, space, cloned) => {
          const name = (local: string) => (prefix ? `${prefix}:${local}` : local);
          const binding = prefix ? `xmlns:${prefix}` : "xmlns";
          const spaceAttributes = {
            absent: "",
            foreign: ' e:space="preserve"',
            default: ' xml:space="default"',
            preserve: ' xml:space="preserve"',
          } satisfies Record<typeof space, string>;
          const spaceAttribute = spaceAttributes[space];
          const before = `<${name("document")} ${binding}="${profile.uri}" xmlns:e="urn:extension"><e:before/><${name("body")}><e:opaque/>`;
          const after = `</${name("body")}><e:after/></${name("document")}>`;
          const xml = `${before}<${name("p")}><${name("r")}><${name("t")}${spaceAttribute}>${whitespace}Text${whitespace}</${name("t")}></${name("r")}></${name("p")}>${after}`;
          const parsed = documentFor(xml, profile.conformance);
          const document = cloned ? structuredClone(parsed) : parsed;
          expect(document.package.document.source?.xml).toBe(xml);
          const saved = serializeTracked(document);
          if (!cloned) {
            expect(saved.startsWith(before)).toBe(true);
            expect(saved.endsWith(after)).toBe(true);
          } else {
            expect(getSourceReplayToken(document)).toBeUndefined();
            // Untracked output derives whitespace metadata from the model,
            // including dropping redundant authored flags around tabs/newlines.
            expect(saved).toBe(serializeDocument(document));
          }
          expect(saved.includes('xml:space="preserve"')).toBe(
            (!cloned && space === "preserve") ||
              requiresXmlSpacePreserve(`${whitespace}Text${whitespace}`),
          );
          if (space === "preserve" && !cloned) expect(saved).toBe(xml);
          const reopened = documentFor(saved, profile.conformance);
          expect(reopened.package.document.content).toEqual(document.package.document.content);
        },
      ),
    );
  });

  test.each(PROFILES)(
    "empty relationship ids require repair inside blocks and the shell in $conformance",
    (profile) => {
      const relationshipUri =
        profile.conformance === DOCX_CONFORMANCE_CLASSES.STRICT
          ? "http://purl.oclc.org/ooxml/officeDocument/relationships"
          : "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
      const prefix = `<x:document xmlns:x="${profile.uri}" xmlns:rel="${relationshipUri}"><x:body>`;
      const paragraph =
        '<x:p><x:hyperlink rel:id=""><x:r><x:t>Text</x:t></x:r></x:hyperlink></x:p>';
      const suffix = "</x:body></x:document>";
      const blockXml = prefix + paragraph + suffix;
      const blockDocument = documentFor(blockXml, profile.conformance);
      expect(blockDocument.package.document.source?.xml).toBe(blockXml);
      const blockSaved = serializeTracked(blockDocument);
      expect(blockSaved.startsWith(prefix)).toBe(true);
      expect(blockSaved.endsWith(suffix)).toBe(true);
      expect(blockSaved).not.toContain('rel:id=""');
      expect(blockSaved).toContain("Text");

      const shellXml = `${prefix}<x:p/><x:sectPr><x:headerReference x:type="default" rel:id=""/></x:sectPr>${suffix}`;
      const shellDocument = documentFor(shellXml, profile.conformance);
      expect(shellDocument.package.document.source?.xml).toBe(shellXml);
      expect(serializeTracked(shellDocument)).not.toContain('rel:id=""');
      expect(serializeTracked(shellDocument)).not.toContain('r:id=""');
    },
  );

  test.each(["original", "cloned"] as const)(
    "caller-replaced source metadata grants no replay authority in a %s capture",
    (owner) => {
      const xml = `<w:document xmlns:w="${PROFILES[0].uri}"><w:body><w:p/></w:body></w:document>`;
      const parsed = documentFor(xml, DOCX_CONFORMANCE_CLASSES.TRANSITIONAL);
      const document = owner === "cloned" ? structuredClone(parsed) : parsed;
      const source = document.package.document.source;
      if (!source) throw new TypeError("Expected synthetic source capture");
      const initial = serializeTracked(document);
      expect(serializeTracked(document)).toBe(initial);
      source.xml = `<!DOCTYPE w:document [<!ENTITY external SYSTEM "file:///unavailable">]>${xml}`;
      expect(serializeTracked(document)).toBe(initial);
      expect(serializeTracked(document)).not.toContain("DOCTYPE");
    },
  );

  test.each(["original", "cloned"] as const)(
    "tracked mutation is rejected while untracked clones rebuild in a %s capture",
    (owner) => {
      const xml = `<w:document xmlns:w="${PROFILES[0].uri}"><w:body><w:p><w:r><w:t>Before</w:t></w:r></w:p></w:body></w:document>`;
      const parsed = documentFor(xml, DOCX_CONFORMANCE_CLASSES.TRANSITIONAL);
      const document = owner === "cloned" ? structuredClone(parsed) : parsed;
      const before = serializeTracked(document);
      const paragraph = document.package.document.content.at(0);
      if (paragraph?.type !== "paragraph") throw new TypeError("Expected synthetic paragraph");
      const mutate = () =>
        paragraph.content.push({ type: "run", content: [{ type: "text", text: " After" }] });
      if (owner === "original") {
        expect(mutate).toThrow();
        expect(serializeTracked(document)).toBe(before);
        return;
      }
      expect(getSourceReplayToken(document)).toBeUndefined();
      mutate();
      const saved = serializeTracked(document);
      expect(saved).toContain(" After");
      expect(saved).not.toBe(xml);
    },
  );

  test.each(["original", "cloned"] as const)(
    "mutation before the first save is retained in a %s capture",
    (owner) => {
      const xml = `<w:document xmlns:w="${PROFILES[0].uri}"><w:body><w:p><w:r><w:t>Before</w:t></w:r></w:p></w:body></w:document>`;
      const parsed = documentFor(xml, DOCX_CONFORMANCE_CLASSES.TRANSITIONAL, "untracked");
      const document = owner === "cloned" ? structuredClone(parsed) : parsed;
      expect(getSourceReplayToken(document)).toBeUndefined();
      const paragraph = document.package.document.content.at(0);
      if (paragraph?.type !== "paragraph") throw new TypeError("Expected synthetic paragraph");
      paragraph.content.push({
        type: "run",
        content: [{ type: "text", text: " Before first save" }],
      });

      const saved = serializeTracked(document);
      expect(saved).toContain("Before first save");
      expect(saved).not.toBe(xml);
    },
  );

  test("untracked edits retain Strict percentage widths through the model serializer", () => {
    const xml = `<x:document xmlns:x="${PROFILES[1].uri}"><x:body><x:tbl><x:tblPr><x:tblW x:w="50%" x:type="pct"/></x:tblPr><x:tblGrid><x:gridCol x:w="2400"/></x:tblGrid><x:tr><x:tc><x:tcPr><x:tcW x:w="50%" x:type="pct"/></x:tcPr><x:p><x:r><x:t>Cell</x:t></x:r></x:p></x:tc></x:tr></x:tbl></x:body></x:document>`;
    const document = documentFor(xml, DOCX_CONFORMANCE_CLASSES.STRICT, "untracked");
    const table = document.package.document.content.at(0);
    if (table?.type !== "table") throw new TypeError("Expected synthetic table");
    const originalWidth = structuredClone(table.formatting?.width);
    const originalCellWidth = structuredClone(table.rows.at(0)?.cells.at(0)?.formatting?.width);
    expect(originalWidth).toEqual({ type: "pct", value: 2500 });
    expect(originalCellWidth).toEqual({ type: "pct", value: 2500 });
    const paragraph = table.rows.at(0)?.cells.at(0)?.content.at(0);
    if (paragraph?.type !== "paragraph") throw new TypeError("Expected synthetic cell paragraph");
    paragraph.content.push({ type: "run", content: [{ type: "text", text: " edited" }] });
    const saved = serializeTracked(document);
    const reopened = parseDocumentBody(saved).content.at(0);
    if (reopened?.type !== "table") throw new TypeError("Expected reopened table");
    expect(reopened.formatting?.width).toEqual(originalWidth);
    expect(reopened.rows.at(0)?.cells.at(0)?.formatting?.width).toEqual(originalCellWidth);
    const reopenedParagraph = reopened.rows.at(0)?.cells.at(0)?.content.at(0);
    if (reopenedParagraph?.type !== "paragraph")
      throw new TypeError("Expected reopened cell paragraph");
    expect(paragraphLogicalText(reopenedParagraph)).toBe(paragraphLogicalText(paragraph));
  });
});

test("unchanged replay serializes zero blocks; one immutable editor edit serializes only its changed block", () => {
  const xml = `<w:document xmlns:w="${PROFILES[0].uri}"><w:body><w:p><w:r><w:t>First</w:t></w:r></w:p><w:p><w:r><w:t>Second</w:t></w:r></w:p></w:body></w:document>`;
  const original = documentFor(xml, DOCX_CONFORMANCE_CLASSES.TRANSITIONAL);
  const captures: Document["package"]["document"]["content"][] = [];
  const serialize = (blocks: Document["package"]["document"]["content"]) => {
    captures.push(blocks);
    return "<w:p><w:r><w:t>Edited</w:t></w:r></w:p>";
  };
  const token = getSourceReplayToken(original);
  expect(replayDocumentSource({ document: original, token, serialize })).toBe(xml);
  expect(replayDocumentSource({ document: original, token, serialize })).toBe(xml);
  expect(captures).toHaveLength(0);
  const projected = toProseDoc(original);
  const state = EditorState.create({ doc: projected });
  const edited = fromProseDoc(
    state.tr.insertText("!", projected.child(0).nodeSize + 1).doc,
    original,
  );
  const changed = edited.package.document.content.at(1);
  expect(edited.package.document.content.at(0)).toBe(original.package.document.content.at(0));
  expect(changed).not.toBe(original.package.document.content.at(1));
  expect(
    replayDocumentSource({ document: edited, token: getSourceReplayToken(edited), serialize }),
  ).toContain("Edited");
  expect(captures).toEqual([[changed]]);
  expect(captures).toHaveLength(1);
  const clone = structuredClone(original);
  expect(replayDocumentSource({ document: clone, token, serialize })).toBeNull();
  expect(captures).toHaveLength(1);
});

test("source-part validation is deferred until output and cached for unchanged bytes", () => {
  const validate = verbatimCapture.isSafeCapturedXmlDocument;
  let validations = 0;
  const spy = spyOn(verbatimCapture, "isSafeCapturedXmlDocument").mockImplementation((xml) => {
    validations += 1;
    return validate(xml);
  });
  try {
    const xml = `<w:document xmlns:w="${PROFILES[0].uri}"><w:body><w:p><w:r><w:t>Source</w:t></w:r></w:p></w:body></w:document>`;
    const document = documentFor(xml, DOCX_CONFORMANCE_CLASSES.TRANSITIONAL);
    expect(validations).toBe(0);
    expect(serializeTracked(document)).toBe(xml);
    expect(validations).toBe(1);
    expect(serializeTracked(document)).toBe(xml);
    expect(validations).toBe(1);

    const cloned = structuredClone(document);
    const regenerated = serializeTracked(cloned);
    expect(getSourceReplayToken(cloned)).toBeUndefined();
    expect(validations).toBe(1);
    expect(serializeTracked(cloned)).toBe(regenerated);
    expect(validations).toBe(1);
  } finally {
    spy.mockRestore();
  }
});
