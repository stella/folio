import { expect, test } from "bun:test";
import JSZip from "jszip";
import { parseDocx } from "./parser";
import { createDocx, repackDocx } from "./rezip";
import { readRootNamespaceBindings } from "./serializer/partNamespaces";
import { createSimpleDocument, serializeDocument } from "./serializer/documentSerializer";
import { captureDocumentSourceBaseline, getDocumentSourceBaseline } from "./headerFooterVerbatim";
import { cloneDocumentWithParagraphPropertySources } from "./paragraphPropertySource";

const WORD = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const untouched =
  '<q:p xmlns:p14="http://schemas.microsoft.com/office/word/2010/wordml" p14:paraId="11111111">\n<q:r>  <q:t>untouched</q:t> </q:r></q:p>';
const source = `<?xml version='1.0' encoding='UTF-8'?>\n<q:document xmlns:q='${WORD}'><q:background q:color='FFFFFF'/>\n<q:body>\n${untouched}\n<!-- authored gap -->\t<q:p><q:r><q:t>before</q:t></q:r></q:p>\n<q:sectPr><q:pgSz q:w='12240' q:h='15840'/></q:sectPr>\n</q:body></q:document>`;

const parsedSource = async () => {
  const zip = await JSZip.loadAsync(await createDocx(createSimpleDocument([{ text: "seed" }])));
  zip.file("word/document.xml", source);
  return parseDocx(await zip.generateAsync({ type: "arraybuffer" }));
};

test("trusted body replay preserves authored shell, metadata, gaps and untouched blocks after cloning", async () => {
  const document = cloneDocumentWithParagraphPropertySources(await parsedSource());
  const paragraph = document.package.document.content.at(1);
  if (paragraph?.type !== "paragraph") throw new Error("Missing paragraph");
  paragraph.content = [{ type: "run", content: [{ type: "text", text: "after" }] }];
  const saved = await repackDocx(document, { bodyAuthority: "canonical" });
  const xml = await (await JSZip.loadAsync(saved)).file("word/document.xml")?.async("text");
  expect(xml).toContain(untouched);
  expect(xml).toContain("<!-- authored gap -->\t");
  expect(xml).toStartWith(`<?xml version='1.0' encoding='UTF-8'?>\n<q:document xmlns:q='${WORD}'>`);
  expect(xml).toContain("<q:background q:color='FFFFFF'/>");
  expect(xml).toContain("<q:sectPr><q:pgSz q:w='12240' q:h='15840'/></q:sectPr>");
  const reopened = await parseDocx(saved);
  const changed = reopened.package.document.content.at(1);
  expect(changed?.type === "paragraph" && changed.content.at(0)).toMatchObject({
    type: "run",
    content: [{ type: "text", text: "after" }],
  });
});

test("document metadata changes serialize from the current model without losing body source", async () => {
  const document = await parsedSource();
  document.package.document.background = { color: { rgb: "ABCDEF" } };
  document.package.document.finalSectionProperties = undefined;
  const xml = serializeDocument(document, undefined, { bodyAuthority: "canonical" });
  expect(xml).toContain(untouched);
  expect(xml).toContain('w:color="ABCDEF"');
  expect(xml).not.toContain("<q:sectPr>");
});

test("trusted source correspondence failures emit a body diagnostic", async () => {
  const document = await parsedSource();
  const diagnostics: unknown[] = [];
  serializeDocument(document, undefined, {
    bodyAuthority: "canonical",
    xml: source + " ",
    onDiagnostic: (value) => diagnostics.push(value),
  });
  expect(diagnostics).toEqual([{ type: "sourceReplayMismatch", part: "word/document.xml" }]);
});

test("unchanged trusted bodies replay exact source across clones and key orders", async () => {
  for (const bufferPresence of ["present", "absent"] as const) {
    const parsed = await parsedSource();
    if (bufferPresence === "absent") delete parsed.originalBuffer;
    for (const document of [parsed, cloneDocumentWithParagraphPropertySources(parsed)]) {
      expect(getDocumentSourceBaseline(document).type).toBe("captured");
      expect(serializeDocument(document, undefined, { bodyAuthority: "canonical" })).toBe(source);
      const { content, ...metadata } = document.package.document;
      document.package.document = { ...metadata, content };
      expect(serializeDocument(document, undefined, { bodyAuthority: "canonical" })).toBe(source);
    }
  }
});

test("trusted clones retain invalid body capture handles instead of recovering package fallback", async () => {
  for (const bufferPresence of ["present", "absent"] as const) {
    for (const invalid of [null, {}, 42]) {
      const document = await parsedSource();
      if (bufferPresence === "absent") delete document.originalBuffer;
      const handle = Object.getOwnPropertySymbols(document.package.document).at(0);
      if (!handle) throw new Error("Missing private body capture handle");
      Object.defineProperty(document.package.document, handle, {
        value: invalid,
        enumerable: true,
        configurable: true,
      });
      expect(getDocumentSourceBaseline(document)).toEqual({ type: "mismatch" });
      const cloned = cloneDocumentWithParagraphPropertySources(document);
      expect(getDocumentSourceBaseline(cloned)).toEqual({ type: "mismatch" });
      const diagnostics: unknown[] = [];
      expect(
        serializeDocument(cloned, undefined, {
          bodyAuthority: "canonical",
          onDiagnostic: (value) => diagnostics.push(value),
        }),
      ).not.toBe(source);
      expect(diagnostics).toEqual([{ type: "sourceReplayMismatch", part: "word/document.xml" }]);
    }
  }
});

test("body replay rejects a modified captured baseline even when current content is unchanged", async () => {
  const document = await parsedSource();
  expect(serializeDocument(document, undefined, { bodyAuthority: "canonical" })).toBe(source);
  const baseline = getDocumentSourceBaseline(document);
  if (baseline.type !== "captured") throw new Error("Missing source baseline");
  baseline.body.content.pop();
  const diagnostics: unknown[] = [];
  const xml = serializeDocument(document, undefined, {
    bodyAuthority: "canonical",
    onDiagnostic: (value) => diagnostics.push(value),
  });
  expect(xml).not.toBe(source);
  expect(diagnostics).toEqual([{ type: "sourceReplayMismatch", part: "word/document.xml" }]);
});

test("unchanged replay retains the block and document cardinality refusal checks", async () => {
  for (const xml of [
    source.replace("</q:document>", "<q:body/></q:document>"),
    source.replace("<q:body>", "<q:background/><q:body>"),
    source.replace("</q:body>", "<q:p/></q:body>"),
  ]) {
    const document = await parsedSource();
    captureDocumentSourceBaseline(document, xml);
    const diagnostics: unknown[] = [];
    expect(
      serializeDocument(document, undefined, {
        bodyAuthority: "canonical",
        onDiagnostic: (value) => diagnostics.push(value),
      }),
    ).not.toBe(xml);
    expect(diagnostics).toEqual([{ type: "sourceReplayMismatch", part: "word/document.xml" }]);
  }
});

test("model authority serializes independently of captured canonical replay state", async () => {
  const document = await parsedSource();
  const diagnostics: unknown[] = [];
  const modelXml = serializeDocument(document, undefined, {
    bodyAuthority: "model",
    onDiagnostic: (value) => diagnostics.push(value),
  });
  expect(modelXml).not.toBe(source);
  expect(serializeDocument(document)).toBe(modelXml);
  const baseline = getDocumentSourceBaseline(document);
  if (baseline.type !== "captured") throw new Error("Missing source baseline");
  baseline.body.content.pop();
  expect(
    serializeDocument(document, undefined, {
      onDiagnostic: (value) => diagnostics.push(value),
    }),
  ).toBe(modelXml);
  const saved = await repackDocx(document);
  const xml = await (await JSZip.loadAsync(saved)).file("word/document.xml")?.async("text");
  expect(xml).toBe(serializeDocument(document, readRootNamespaceBindings(source)));
  expect(diagnostics).toEqual([]);
});
