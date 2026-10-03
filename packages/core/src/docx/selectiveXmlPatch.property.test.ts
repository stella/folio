/**
 * Property tests for `buildPatchedDocumentXml` and helpers.
 *
 * The selective patcher mutates raw OOXML strings. We use fast-check to fuzz
 * the structural invariants that the patcher must preserve regardless of the
 * actual paragraph content:
 *
 *   1. An empty change set is the identity.
 *   2. Patching with the same serialized XML is the identity.
 *   3. Patching one paragraph only mutates that paragraph's slice.
 *   4. Patch order does not matter (the user-visible result must not depend
 *      on Set iteration order).
 *   5. Paragraph count must be invariant across patch.
 *   6. `findParagraphOffsets` returns a well-formed span.
 *   7. Nested paragraph changes cannot splice overlapping source ranges.
 */

import { describe, setDefaultTimeout, test, expect } from "bun:test";
import fc from "fast-check";
import JSZip from "jszip";
import { validateDocxPackage } from "@stll/docx-core";
import { createEmptyDocument } from "../utils/createDocument";
import { createDocx, repackDocx } from "./rezip";
import { parseDocx } from "./parser";
import { attemptSelectiveSave } from "./selectiveSave";

import {
  assertProperty,
  propertyConfig,
  propertyTestTimeout,
} from "../../../../test/property-testing";

import {
  buildPatchedDocumentXml,
  countParagraphElements,
  findParagraphOffsets,
} from "./selectiveXmlPatch";
import { buildStructuralDocumentPatch } from "./structuralXmlPatch";
import { parseXmlDocument } from "./xmlParser";
import { serializeResolutionJoins } from "./reviewResolutionProvenance";
import { FOLIO_REVIEW_HISTORY_NAMESPACE } from "./reviewHistoryNamespace";

setDefaultTimeout(propertyTestTimeout(30_000));

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

const paraIdArb = fc.stringMatching(/^[A-Z]{3}[0-9]{4}$/u).map((id) => id.toUpperCase());

const paraTextArb = fc.stringMatching(/^[A-Za-z0-9 ]{1,40}$/u);

type RunMark = "b" | "i" | "u";

type Para = { id: string; text: string; marks: RunMark[]; list: boolean };

// Emphasis and list numbering the patcher must carry through untouched. `<w:pPr>`
// and `<w:rPr>` both start with the bytes `<w:p`/`<w:r`, so they exercise the
// paragraph-counter and offset scanner's guards against mistaking a property
// element for a real `<w:p>`/`<w:r>` boundary.
const marksArb = fc.subarray(["b", "i", "u"] as RunMark[], { minLength: 0, maxLength: 3 });

const paragraphsArb: fc.Arbitrary<Para[]> = fc.uniqueArray(
  fc.record({ id: paraIdArb, text: paraTextArb, marks: marksArb, list: fc.boolean() }),
  {
    selector: (p) => p.id,
    minLength: 2,
    maxLength: 8,
  },
);

function renderRunProps(marks: RunMark[]): string {
  if (marks.length === 0) {
    return "";
  }
  const tags = marks.map((m) => (m === "u" ? '<w:u w:val="single"/>' : `<w:${m}/>`)).join("");
  return `<w:rPr>${tags}</w:rPr>`;
}

function renderParaProps(list: boolean): string {
  if (!list) {
    return "";
  }
  return '<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>';
}

function renderParagraph(p: Para): string {
  return `<w:p w14:paraId="${p.id}">${renderParaProps(p.list)}<w:r>${renderRunProps(p.marks)}<w:t>${p.text}</w:t></w:r></w:p>`;
}

function renderDoc(paras: Para[]): string {
  return `${XML_DECL}
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
<w:body>
${paras.map(renderParagraph).join("\n")}
<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>
</w:body>
</w:document>`;
}

const REVIEW_JOIN_ATTRIBUTES = serializeResolutionJoins({ before: 1, after: 1, remove: 0 });

type ReviewDocShape =
  | "plain"
  | "candidate"
  | "supported"
  | "used-without-ignorable"
  | "supported-alias"
  | "conflicting"
  | "bound-without-ignorable";

const renderReviewDoc = (text: string, shape: ReviewDocShape) => {
  const folioNamespace = shape === "conflicting" ? "urn:other" : FOLIO_REVIEW_HISTORY_NAMESPACE;
  const review =
    shape === "candidate" || shape === "supported" || shape === "used-without-ignorable";
  const ignorable =
    shape === "candidate" ||
    shape === "supported" ||
    shape === "supported-alias" ||
    shape === "conflicting"
      ? ' mc:Ignorable="folio"'
      : "";
  const alias =
    shape === "supported-alias" ? ` xmlns:history="${FOLIO_REVIEW_HISTORY_NAMESPACE}"` : "";
  const ignorablePrefix = shape === "supported-alias" ? "history" : "folio";
  return `${XML_DECL}<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"${shape === "plain" ? "" : ` xmlns:folio="${folioNamespace}"`}${alias}${ignorable ? ` mc:Ignorable="${ignorablePrefix}"` : ""}><w:body><w:p w14:paraId="A0000001">${review ? `<w:ins w:id="1" w:author="Reviewer"${REVIEW_JOIN_ATTRIBUTES}>` : ""}<w:r><w:t>${text}</w:t></w:r>${review ? "</w:ins>" : ""}</w:p><w:p w14:paraId="A0000002"><w:r><w:t>untouched</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>`;
};

describe("buildPatchedDocumentXml property invariants", () => {
  test("selective save preserves introduced review provenance through valid package reopen", async () => {
    await assertProperty(
      fc.asyncProperty(
        paraTextArb,
        fc.constantFrom("paragraph", "structural"),
        async (text, route) => {
          const initial = createEmptyDocument();
          initial.package.document.content = [
            {
              type: "paragraph",
              paraId: "A0000001",
              content: [{ type: "run", content: [{ type: "text", text: "before" }] }],
            },
          ];
          const source = await createDocx(initial);
          const zip = await JSZip.loadAsync(source);
          expect(await zip.file("word/document.xml")?.async("text")).not.toContain("xmlns:folio=");
          const document = await parseDocx(source, { preloadFonts: false, detectVariables: false });
          const paragraph = document.package.document.content.at(0);
          if (paragraph?.type !== "paragraph") throw new TypeError("Expected source paragraph");
          const joins = { before: 1, after: 1, remove: 0 };
          paragraph.content = [
            {
              type: "insertion",
              info: { id: 1, author: "Reviewer" },
              resolutionJoins: joins,
              content: [{ type: "run", content: [{ type: "text", text }] }],
            },
          ];
          const selective = await attemptSelectiveSave(document, source, {
            changedParaIds: new Set(["A0000001"]),
            structuralChange: route === "structural",
            hasUntrackedChanges: false,
          });
          if (route === "paragraph") expect(selective).toBeNull();
          // Structural save may also decline for package-wide invariants;
          // either route must reopen with the complete review provenance.
          const saved = selective ?? (await repackDocx(document, { updateModifiedDate: false }));
          expect(await validateDocxPackage(new Uint8Array(saved))).toEqual({ valid: true });
          const reopened = await parseDocx(saved, { preloadFonts: false, detectVariables: false });
          const restored = reopened.package.document.content.at(0);
          if (restored?.type !== "paragraph") throw new TypeError("Expected reopened paragraph");
          const insertion = restored.content.find((item) => item.type === "insertion");
          expect(insertion?.resolutionJoins).toEqual(joins);
        },
      ),
      {
        numRuns: 4,
        examples: [
          ["a", "paragraph"],
          ["a", "structural"],
        ],
      },
    );
  });

  test("introduced review namespace is refused by ordinary splices and locally bound by structural splices", () => {
    assertProperty(
      fc.property(paraTextArb, (text) => {
        const original = renderReviewDoc("before", "plain");
        const serialized = renderReviewDoc(text, "candidate");
        const changed = new Set(["A0000001"]);

        expect(buildPatchedDocumentXml(original, serialized, changed)).toBeNull();

        const structural = buildStructuralDocumentPatch({
          originalXml: original,
          serializedXml: serialized,
          changedIds: changed,
        });
        expect(structural).not.toBeNull();
        expect(structural).toContain(`xmlns:folio="${FOLIO_REVIEW_HISTORY_NAMESPACE}"`);
        expect(structural).toMatch(/mc:Ignorable="[^"]*folio/u);
        expect(parseXmlDocument(structural ?? "")).not.toBeNull();
      }),
      { numRuns: 20 },
    );
  });

  test("ordinary and structural splices preserve supported review namespace bindings", () => {
    assertProperty(
      fc.property(paraTextArb, (text) => {
        const original = renderReviewDoc("before", "supported");
        const serialized = renderReviewDoc(text, "supported");
        const changed = new Set(["A0000001"]);

        expect(buildPatchedDocumentXml(original, serialized, changed)).not.toBeNull();
        expect(
          buildStructuralDocumentPatch({
            originalXml: original,
            serializedXml: serialized,
            changedIds: changed,
          }),
        ).not.toBeNull();
      }),
      { numRuns: 20 },
    );
  });

  test("existing extension uses retain source roots that omit ignorable metadata", () => {
    assertProperty(
      fc.property(paraTextArb, (text) => {
        const original = renderReviewDoc("before", "used-without-ignorable");
        const serialized = renderReviewDoc(text, "candidate");
        const patched = buildPatchedDocumentXml(original, serialized, new Set(["A0000001"]));
        expect(patched).not.toBeNull();
        expect(patched).not.toContain("mc:Ignorable=");
        expect(patched).toContain(REVIEW_JOIN_ATTRIBUTES);
      }),
      { numRuns: 20 },
    );
  });

  test("ignorable namespace support is compared by URI across aliases", () => {
    assertProperty(
      fc.property(paraTextArb, (text) => {
        const original = renderReviewDoc("before", "supported-alias");
        const serialized = renderReviewDoc(text, "candidate");
        expect(buildPatchedDocumentXml(original, serialized, new Set(["A0000001"]))).not.toBeNull();
      }),
      { numRuns: 20 },
    );
  });

  test("conflicting or non-ignorable root bindings require the full serializer", () => {
    assertProperty(
      fc.property(
        paraTextArb,
        fc.constantFrom("conflicting", "bound-without-ignorable"),
        (text, sourceShape) => {
          const original = renderReviewDoc("before", sourceShape);
          const serialized = renderReviewDoc(text, "candidate");
          expect(buildPatchedDocumentXml(original, serialized, new Set(["A0000001"]))).toBeNull();
          expect(
            buildStructuralDocumentPatch({
              originalXml: original,
              serializedXml: serialized,
              changedIds: new Set(["A0000001"]),
            }),
          ).not.toBeNull();
        },
      ),
      { numRuns: 20 },
    );
  });

  test("local ignorable metadata on one replacement does not license a sibling", () => {
    assertProperty(
      fc.property(paraTextArb, paraTextArb, (first, second) => {
        const original = renderReviewDoc("before", "bound-without-ignorable");
        const serialized = `${XML_DECL}<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:folio="${FOLIO_REVIEW_HISTORY_NAMESPACE}" mc:Ignorable="folio"><w:body><w:p w14:paraId="A0000001" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="folio"><w:ins w:id="1" w:author="Reviewer"${REVIEW_JOIN_ATTRIBUTES}><w:r><w:t>${first}</w:t></w:r></w:ins></w:p><w:p w14:paraId="A0000002"><w:ins w:id="2" w:author="Reviewer"${REVIEW_JOIN_ATTRIBUTES}><w:r><w:t>${second}</w:t></w:r></w:ins></w:p><w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>`;
        expect(
          buildPatchedDocumentXml(original, serialized, new Set(["A0000001", "A0000002"])),
        ).toBeNull();
      }),
      { numRuns: 20 },
    );
  });

  test("empty change set returns the original XML unchanged", () => {
    fc.assert(
      fc.property(paragraphsArb, (paras) => {
        const xml = renderDoc(paras);
        const result = buildPatchedDocumentXml(xml, xml, new Set());
        expect(result).toBe(xml);
      }),
      propertyConfig({ numRuns: 50 }),
    );
  });

  test("patching against identical serialized XML is the identity", () => {
    fc.assert(
      fc.property(paragraphsArb, (paras) => {
        const xml = renderDoc(paras);
        const ids = new Set(paras.map((p) => p.id));
        const result = buildPatchedDocumentXml(xml, xml, ids);
        expect(result).toBe(xml);
      }),
      propertyConfig({ numRuns: 50 }),
    );
  });

  test("nested paragraph edits never apply overlapping source ranges", () => {
    assertProperty(
      fc.property(paraTextArb, paraTextArb, fc.boolean(), (before, after, reversed) => {
        const renderNested = (text: string) =>
          `${XML_DECL}<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body>` +
          `<w:p w14:paraId="A0000001"><w:r><w:pict><w:txbxContent>` +
          renderParagraph({ id: "A0000002", text, marks: [], list: false }) +
          `</w:txbxContent></w:pict></w:r></w:p>` +
          renderParagraph({ id: "A0000003", text: "Tail", marks: [], list: false }) +
          `</w:body></w:document>`;
        const ids = reversed ? ["A0000002", "A0000001"] : ["A0000001", "A0000002"];
        expect(
          buildPatchedDocumentXml(renderNested(before), renderNested(after), new Set(ids)),
        ).toBeNull();
      }),
      { numRuns: 50 },
    );
  });

  test("paragraph count is preserved after a patch", () => {
    fc.assert(
      fc.property(paragraphsArb, fc.integer({ min: 0 }), (paras, seed) => {
        const idx = seed % paras.length;
        const original = renderDoc(paras);
        const edited = paras.map((p, i) => (i === idx ? { ...p, text: `${p.text}_NEW` } : p));
        const serialized = renderDoc(edited);
        const ids = new Set([edited[idx]!.id]);
        const result = buildPatchedDocumentXml(original, serialized, ids);
        expect(result).not.toBeNull();
        if (result) {
          expect(countParagraphElements(result)).toBe(countParagraphElements(original));
        }
      }),
      propertyConfig({ numRuns: 50 }),
    );
  });

  test("bytes outside the changed paragraph remain byte-identical", () => {
    fc.assert(
      fc.property(paragraphsArb, fc.integer({ min: 0 }), (paras, seed) => {
        const idx = seed % paras.length;
        const target = paras[idx]!;
        const original = renderDoc(paras);
        const edited = paras.map((p, i) => (i === idx ? { ...p, text: `${p.text}_X` } : p));
        const serialized = renderDoc(edited);
        const result = buildPatchedDocumentXml(original, serialized, new Set([target.id]));
        if (!result) {
          return;
        }

        const originalOffsets = findParagraphOffsets(original, target.id);
        const resultOffsets = findParagraphOffsets(result, target.id);
        expect(originalOffsets).not.toBeNull();
        expect(resultOffsets).not.toBeNull();
        if (originalOffsets && resultOffsets) {
          expect(result.slice(0, resultOffsets.start)).toBe(
            original.slice(0, originalOffsets.start),
          );
          expect(result.slice(resultOffsets.end)).toBe(original.slice(originalOffsets.end));
        }
      }),
      propertyConfig({ numRuns: 50 }),
    );
  });

  test("patch result is independent of Set iteration order over changed ids", () => {
    fc.assert(
      fc.property(paragraphsArb, (paras) => {
        if (paras.length < 2) {
          return;
        }
        const original = renderDoc(paras);
        const edited = paras.map((p, i) => (i < 2 ? { ...p, text: `${p.text}_X` } : p));
        const serialized = renderDoc(edited);

        const orderA = new Set([paras[0]!.id, paras[1]!.id]);
        const orderB = new Set([paras[1]!.id, paras[0]!.id]);
        const a = buildPatchedDocumentXml(original, serialized, orderA);
        const b = buildPatchedDocumentXml(original, serialized, orderB);
        expect(a).toBe(b);
      }),
      propertyConfig({ numRuns: 50 }),
    );
  });

  test("findParagraphOffsets returns a well-formed span", () => {
    fc.assert(
      fc.property(paragraphsArb, (paras) => {
        const xml = renderDoc(paras);
        for (const p of paras) {
          const offsets = findParagraphOffsets(xml, p.id);
          expect(offsets).not.toBeNull();
          if (offsets) {
            expect(offsets.end).toBeGreaterThan(offsets.start);
            const slice = xml.slice(offsets.start, offsets.end);
            expect(slice.startsWith("<w:p")).toBe(true);
            expect(slice.endsWith("</w:p>")).toBe(true);
          }
        }
      }),
      propertyConfig({ numRuns: 50 }),
    );
  });
});
