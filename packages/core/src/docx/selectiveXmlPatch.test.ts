/**
 * Unit tests for selectiveXmlPatch
 */

import { describe, test, expect } from "bun:test";

import {
  findParagraphOffsets,
  extractParagraphXml,
  buildParagraphOffsetIndex,
  validatePatchSafety,
  buildPatchedDocumentXml,
  buildPatchedNotePartXml,
  countParagraphElements,
  type NotePartPatch,
} from "./selectiveXmlPatch";

const patchedXmlOf = (patch: NotePartPatch): string => {
  if (patch.type !== "patched") {
    throw new Error(`note part patch refused: ${patch.reason}`);
  }
  return patch.xml;
};

// ============================================================================
// Test XML fixtures
// ============================================================================

const SIMPLE_DOC = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
<w:body>
<w:p w14:paraId="AAA111" w14:textId="T1"><w:pPr><w:jc w:val="left"/></w:pPr><w:r><w:t>First paragraph</w:t></w:r></w:p>
<w:p w14:paraId="BBB222" w14:textId="T2"><w:r><w:t>Second paragraph</w:t></w:r></w:p>
<w:p w14:paraId="CCC333" w14:textId="T3"><w:r><w:t>Third paragraph</w:t></w:r></w:p>
</w:body>
</w:document>`;

const DOC_WITH_MC = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
<w:body>
<w:p w14:paraId="OUTER1" w14:textId="T1"><mc:AlternateContent><mc:Choice Requires="wps"><w:p w14:paraId="INNER1"><w:r><w:t>Inner</w:t></w:r></w:p></mc:Choice><mc:Fallback><w:p w14:paraId="INNER2"><w:r><w:t>Fallback</w:t></w:r></w:p></mc:Fallback></mc:AlternateContent></w:p>
<w:p w14:paraId="NORMAL1" w14:textId="T2"><w:r><w:t>Normal paragraph</w:t></w:r></w:p>
</w:body>
</w:document>`;

const DOC_WITH_DUPLICATE_ID = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
<w:body>
<w:p w14:paraId="DUP001" w14:textId="T1"><w:r><w:t>First</w:t></w:r></w:p>
<w:p w14:paraId="DUP001" w14:textId="T2"><w:r><w:t>Duplicate</w:t></w:r></w:p>
</w:body>
</w:document>`;

const DOC_WITH_MANY_ATTRS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
<w:body>
<w:p w:rsidR="00A12345" w:rsidRDefault="00B67890" w14:paraId="ATTR01" w14:textId="TXID01" w:rsidP="00C11111"><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Heading</w:t></w:r></w:p>
</w:body>
</w:document>`;

// ============================================================================
// findParagraphOffsets
// ============================================================================

describe("findParagraphOffsets", () => {
  test("finds a simple paragraph by paraId", () => {
    const offsets = findParagraphOffsets(SIMPLE_DOC, "AAA111");
    expect(offsets).not.toBeNull();
    if (!offsets) {
      throw new Error("Expected offsets");
    }
    const extracted = SIMPLE_DOC.slice(offsets.start, offsets.end);
    expect(extracted).toStartWith('<w:p w14:paraId="AAA111"');
    expect(extracted).toEndWith("</w:p>");
    expect(extracted).toContain("First paragraph");
  });

  test("finds the second paragraph", () => {
    const offsets = findParagraphOffsets(SIMPLE_DOC, "BBB222");
    expect(offsets).not.toBeNull();
    if (!offsets) {
      throw new Error("Expected offsets");
    }
    const extracted = SIMPLE_DOC.slice(offsets.start, offsets.end);
    expect(extracted).toContain("Second paragraph");
  });

  test("finds the third paragraph", () => {
    const offsets = findParagraphOffsets(SIMPLE_DOC, "CCC333");
    expect(offsets).not.toBeNull();
    if (!offsets) {
      throw new Error("Expected offsets");
    }
    const extracted = SIMPLE_DOC.slice(offsets.start, offsets.end);
    expect(extracted).toContain("Third paragraph");
  });

  test("returns null for missing paraId", () => {
    expect(findParagraphOffsets(SIMPLE_DOC, "MISSING")).toBeNull();
  });

  test("returns null for duplicate paraId", () => {
    expect(findParagraphOffsets(DOC_WITH_DUPLICATE_ID, "DUP001")).toBeNull();
  });

  test("handles nested w:p inside mc:AlternateContent", () => {
    // The outer paragraph should encompass all nested content
    const offsets = findParagraphOffsets(DOC_WITH_MC, "OUTER1");
    expect(offsets).not.toBeNull();
    if (!offsets) {
      throw new Error("Expected offsets");
    }
    const extracted = DOC_WITH_MC.slice(offsets.start, offsets.end);
    expect(extracted).toStartWith('<w:p w14:paraId="OUTER1"');
    expect(extracted).toEndWith("</w:p>");
    // Should contain the nested mc:AlternateContent
    expect(extracted).toContain("mc:AlternateContent");
    expect(extracted).toContain("INNER1");
    expect(extracted).toContain("INNER2");
  });

  test("finds normal paragraph after mc:AlternateContent block", () => {
    const offsets = findParagraphOffsets(DOC_WITH_MC, "NORMAL1");
    expect(offsets).not.toBeNull();
    if (!offsets) {
      throw new Error("Expected offsets");
    }
    const extracted = DOC_WITH_MC.slice(offsets.start, offsets.end);
    expect(extracted).toContain("Normal paragraph");
    expect(extracted).not.toContain("mc:AlternateContent");
  });

  test("finds paragraph with many attributes", () => {
    const offsets = findParagraphOffsets(DOC_WITH_MANY_ATTRS, "ATTR01");
    expect(offsets).not.toBeNull();
    if (!offsets) {
      throw new Error("Expected offsets");
    }
    const extracted = DOC_WITH_MANY_ATTRS.slice(offsets.start, offsets.end);
    expect(extracted).toContain("Heading");
    expect(extracted).toContain('w:rsidR="00A12345"');
  });

  test("handles self-closing paragraph tag", () => {
    const xml = '<w:body><w:p w14:paraId="SELF01"/></w:body>';
    const offsets = findParagraphOffsets(xml, "SELF01");
    expect(offsets).not.toBeNull();
    if (!offsets) {
      throw new Error("Expected offsets");
    }
    const extracted = xml.slice(offsets.start, offsets.end);
    expect(extracted).toBe('<w:p w14:paraId="SELF01"/>');
  });
});

// ============================================================================
// extractParagraphXml
// ============================================================================

describe("extractParagraphXml", () => {
  test("extracts a paragraph by paraId", () => {
    const xml = extractParagraphXml(SIMPLE_DOC, "BBB222");
    expect(xml).not.toBeNull();
    expect(xml).toContain("Second paragraph");
    expect(xml).toStartWith("<w:p");
    expect(xml).toEndWith("</w:p>");
  });

  test("returns null for missing paraId", () => {
    expect(extractParagraphXml(SIMPLE_DOC, "NOPE")).toBeNull();
  });
});

// ============================================================================
// buildParagraphOffsetIndex
// ============================================================================

describe("buildParagraphOffsetIndex", () => {
  test("indexes every unique paraId in one pass, matching per-id findParagraphOffsets results", () => {
    // Regression guard: rezip.ts's collectChangedNoteParaIds used to call
    // extractParagraphXml (a full regex-scan-plus-depth-walk) once per note
    // paraId, which is O(note count * XML size). The index below is built
    // once per XML and every lookup afterward is O(1); this test asserts the
    // index agrees with the original per-id function it replaces.
    const index = buildParagraphOffsetIndex(SIMPLE_DOC);
    for (const id of ["AAA111", "BBB222", "CCC333"]) {
      const expected = findParagraphOffsets(SIMPLE_DOC, id);
      expect(index.get(id)).toEqual(expected);
    }
    expect(index.size).toBe(3);
  });

  test("indexes nested <w:p> inside mc:AlternateContent alongside the outer paragraph", () => {
    const index = buildParagraphOffsetIndex(DOC_WITH_MC);
    for (const id of ["OUTER1", "INNER1", "INNER2", "NORMAL1"]) {
      const expected = findParagraphOffsets(DOC_WITH_MC, id);
      expect(index.get(id)).toEqual(expected);
    }
  });

  test("omits a paraId that appears on more than one element, mirroring the ambiguous-null case", () => {
    const index = buildParagraphOffsetIndex(DOC_WITH_DUPLICATE_ID);
    expect(index.has("DUP001")).toBe(false);
    expect(findParagraphOffsets(DOC_WITH_DUPLICATE_ID, "DUP001")).toBeNull();
  });

  test("returns an empty index for a missing paraId lookup", () => {
    const index = buildParagraphOffsetIndex(SIMPLE_DOC);
    expect(index.get("MISSING")).toBeUndefined();
  });
});

// ============================================================================
// countParagraphElements
// ============================================================================

describe("countParagraphElements", () => {
  test("counts paragraphs in simple doc", () => {
    expect(countParagraphElements(SIMPLE_DOC)).toBe(3);
  });

  test("counts all w:p elements including nested ones", () => {
    // OUTER1 + INNER1 + INNER2 + NORMAL1 = 4
    expect(countParagraphElements(DOC_WITH_MC)).toBe(4);
  });

  test("does not count w:pPr or w:pStyle as paragraphs", () => {
    const xml =
      '<w:body><w:p w14:paraId="X"><w:pPr><w:pStyle w:val="Normal"/></w:pPr></w:p></w:body>';
    expect(countParagraphElements(xml)).toBe(1);
  });
});

// ============================================================================
// validatePatchSafety
// ============================================================================

describe("validatePatchSafety", () => {
  test("safe when all changed IDs exist in both XMLs", () => {
    const result = validatePatchSafety(SIMPLE_DOC, SIMPLE_DOC, new Set(["AAA111"]));
    expect(result.safe).toBe(true);
  });

  test("safe with empty changed set", () => {
    const result = validatePatchSafety(SIMPLE_DOC, SIMPLE_DOC, new Set());
    expect(result.safe).toBe(true);
  });

  test("unsafe when paraId not found in original", () => {
    const result = validatePatchSafety(SIMPLE_DOC, SIMPLE_DOC, new Set(["MISSING"]));
    expect(result.safe).toBe(false);
    expect(result.reason).toContain("paraId-not-found-in-original");
  });

  test("unsafe when paraId not found in serialized", () => {
    const serializedWithout = SIMPLE_DOC.replace("AAA111", "NEWID1");
    const result = validatePatchSafety(SIMPLE_DOC, serializedWithout, new Set(["AAA111"]));
    expect(result.safe).toBe(false);
    expect(result.reason).toContain("paraId-not-found-in-serialized");
  });

  test("unsafe when duplicate paraId in original", () => {
    const result = validatePatchSafety(
      DOC_WITH_DUPLICATE_ID,
      DOC_WITH_DUPLICATE_ID,
      new Set(["DUP001"]),
    );
    expect(result.safe).toBe(false);
    expect(result.reason).toContain("duplicate-paraId-in-original");
  });

  test("safe when the serialization holds a paragraph the source lacks elsewhere", () => {
    const serializedExtra = SIMPLE_DOC.replace(
      "</w:body>",
      '<w:p w14:paraId="DDD444"><w:r><w:t>Extra</w:t></w:r></w:p></w:body>',
    );
    const result = validatePatchSafety(SIMPLE_DOC, serializedExtra, new Set(["AAA111"]));
    expect(result.safe).toBe(true);
  });
});

// ============================================================================
// buildPatchedDocumentXml
// ============================================================================

describe("buildPatchedDocumentXml", () => {
  test("returns original XML when no changes", () => {
    const result = buildPatchedDocumentXml(SIMPLE_DOC, SIMPLE_DOC, new Set());
    expect(result).toBe(SIMPLE_DOC);
  });

  test("replaces a single paragraph", () => {
    // Create a "serialized" version where the first paragraph has different text
    const serialized = SIMPLE_DOC.replace("First paragraph", "MODIFIED paragraph");
    const result = buildPatchedDocumentXml(SIMPLE_DOC, serialized, new Set(["AAA111"]));

    expect(result).not.toBeNull();
    // Changed paragraph should have new content
    expect(result).toContain("MODIFIED paragraph");
    // Unchanged paragraphs should be byte-for-byte identical
    expect(result).toContain("Second paragraph");
    expect(result).toContain("Third paragraph");
  });

  test("replaces multiple paragraphs", () => {
    const serialized = SIMPLE_DOC.replace("First paragraph", "MODIFIED first").replace(
      "Third paragraph",
      "MODIFIED third",
    );
    const result = buildPatchedDocumentXml(SIMPLE_DOC, serialized, new Set(["AAA111", "CCC333"]));

    expect(result).not.toBeNull();
    expect(result).toContain("MODIFIED first");
    expect(result).toContain("Second paragraph"); // unchanged
    expect(result).toContain("MODIFIED third");
  });

  test("handles replacement with longer XML", () => {
    const serialized = SIMPLE_DOC.replace(
      "<w:r><w:t>First paragraph</w:t></w:r>",
      "<w:r><w:rPr><w:b/></w:rPr><w:t>Much longer first paragraph with bold formatting and extra content</w:t></w:r>",
    );
    const result = buildPatchedDocumentXml(SIMPLE_DOC, serialized, new Set(["AAA111"]));

    expect(result).not.toBeNull();
    expect(result).toContain("Much longer first paragraph");
    expect(result).toContain("Second paragraph");
  });

  test("handles replacement with shorter XML", () => {
    const serialized = SIMPLE_DOC.replace(
      '<w:pPr><w:jc w:val="left"/></w:pPr><w:r><w:t>First paragraph</w:t></w:r>',
      "<w:r><w:t>Hi</w:t></w:r>",
    );
    const result = buildPatchedDocumentXml(SIMPLE_DOC, serialized, new Set(["AAA111"]));

    expect(result).not.toBeNull();
    expect(result).toContain("Hi");
    expect(result).toContain("Second paragraph");
  });

  test("returns null when paraId missing from original", () => {
    const result = buildPatchedDocumentXml(SIMPLE_DOC, SIMPLE_DOC, new Set(["MISSING"]));
    expect(result).toBeNull();
  });

  test("splices only the changed paragraph when the serialization has an extra one", () => {
    const serializedExtra = SIMPLE_DOC.replace("First paragraph", "MODIFIED first").replace(
      "</w:body>",
      '<w:p w14:paraId="DDD444"><w:r><w:t>Extra</w:t></w:r></w:p></w:body>',
    );
    const result = buildPatchedDocumentXml(SIMPLE_DOC, serializedExtra, new Set(["AAA111"]));
    expect(result).toBe(SIMPLE_DOC.replace("First paragraph", "MODIFIED first"));
  });

  test("preserves bytes around unchanged paragraphs exactly", () => {
    const serialized = SIMPLE_DOC.replace("Second paragraph", "CHANGED second");
    const result = buildPatchedDocumentXml(SIMPLE_DOC, serialized, new Set(["BBB222"]));
    expect(result).not.toBeNull();

    // Extract the part before the changed paragraph — should be identical
    const origBeforeBBB = SIMPLE_DOC.slice(0, SIMPLE_DOC.indexOf('<w:p w14:paraId="BBB222"'));
    if (!result) {
      throw new Error("Expected result");
    }
    const resultBeforeBBB = result.slice(0, result.indexOf('<w:p w14:paraId="BBB222"'));
    expect(resultBeforeBBB).toBe(origBeforeBBB);

    // Extract the part after the changed paragraph
    const origAfterBBB = SIMPLE_DOC.slice(SIMPLE_DOC.indexOf('<w:p w14:paraId="CCC333"'));
    const resultAfterBBB = result.slice(result.indexOf('<w:p w14:paraId="CCC333"'));
    expect(resultAfterBBB).toBe(origAfterBBB);
  });
});

// ============================================================================
// Local routing: each changed paragraph is placed on its own evidence
// ============================================================================

const ROUTING_ROOT =
  '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:v="urn:schemas-microsoft-com:vml">';

const routingDoc = (body: string): string =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${ROUTING_ROOT}<w:body>${body}</w:body></w:document>`;

/** A paragraph with one text run, carrying `id` when one is given. */
const para = (id: string | undefined, text: string): string =>
  `<w:p${id === undefined ? "" : ` w14:paraId="${id}"`}><w:r><w:t>${text}</w:t></w:r></w:p>`;

/** A paragraph whose run holds `runContent` (a text box). */
const hostPara = (id: string | undefined, runContent: string): string =>
  `<w:p${id === undefined ? "" : ` w14:paraId="${id}"`}><w:r>${runContent}</w:r></w:p>`;

const drawingTextBox = (paragraphs: string): string =>
  `<w:drawing><wps:wsp><wps:txbx><w:txbxContent>${paragraphs}</w:txbxContent></wps:txbx></wps:wsp></w:drawing>`;

const vmlTextBox = (paragraphs: string): string =>
  `<w:pict><v:shape><v:textbox><w:txbxContent>${paragraphs}</w:txbxContent></v:textbox></v:shape></w:pict>`;

const alternateContent = (choice: string, fallback: string): string =>
  `<mc:AlternateContent><mc:Choice Requires="wps">${choice}</mc:Choice><mc:Fallback>${fallback}</mc:Fallback></mc:AlternateContent>`;

const tableCell = (content: string): string =>
  `<w:tbl><w:tr><w:tc>${content}</w:tc></w:tr></w:tbl>`;

const refusalOf = (original: string, serialized: string, id: string): string | undefined => {
  const result = validatePatchSafety(original, serialized, new Set([id]));
  expect(result.safe).toBe(false);
  expect(buildPatchedDocumentXml(original, serialized, new Set([id]))).toBeNull();
  return result.reason;
};

describe("selective patch routes each changed paragraph locally", () => {
  test("splices an authored paragraph when the serialization writes no text-box Fallback", () => {
    const box = hostPara(
      "B0000001",
      alternateContent(
        drawingTextBox(para("B0000002", "Box")),
        vmlTextBox(para("B0000002", "Box")),
      ),
    );
    const original = routingDoc(para("A0000001", "Edit me") + box + para("A0000002", "Tail"));
    const serialized = routingDoc(
      para("A0000001", "Edit me!") +
        hostPara("B0000001", drawingTextBox(para("B0000002", "Box"))) +
        para("A0000002", "Tail"),
    );

    expect(buildPatchedDocumentXml(original, serialized, new Set(["A0000001"]))).toBe(
      original.replace("Edit me", "Edit me!"),
    );
  });

  test("splices an id-less paragraph by its main-flow ordinal when a text box reads differently", () => {
    // The source's VML box holds two paragraphs; the model's box holds one.
    const original = routingDoc(
      para(undefined, "One") +
        hostPara(undefined, vmlTextBox(para(undefined, "Box a") + para(undefined, "Box b"))) +
        para(undefined, "Two"),
    );
    const serialized = routingDoc(
      para("0000000A", "One") +
        hostPara("0000000B", drawingTextBox(para("0000000C", "Box a"))) +
        para("0000000D", "Two!"),
    );

    // The minted id names nothing in the file, so it is not written into it.
    expect(buildPatchedDocumentXml(original, serialized, new Set(["0000000D"]))).toBe(
      original.replace("Two", "Two!"),
    );
  });

  test("splices an id-less text-box paragraph when the text boxes line up", () => {
    const original = routingDoc(
      para(undefined, "One") + hostPara(undefined, vmlTextBox(para(undefined, "Box"))),
    );
    const serialized = routingDoc(
      para("0000000A", "One") + hostPara("0000000B", vmlTextBox(para("0000000C", "Box!"))),
    );

    expect(buildPatchedDocumentXml(original, serialized, new Set(["0000000C"]))).toBe(
      original.replace("Box", "Box!"),
    );
  });

  test("refuses an id-less paragraph whose story's ordinals no longer line up", () => {
    const original = routingDoc(para(undefined, "One") + para(undefined, "Two"));
    const serialized = routingDoc(
      para("0000000A", "One") + para("0000000B", "Extra") + para("0000000C", "Two!"),
    );

    expect(refusalOf(original, serialized, "0000000C")).toBe(
      "unaligned-paragraph-ordinals: 0000000C",
    );
  });

  test("refuses an id-less text-box paragraph when the text boxes disagree", () => {
    const original = routingDoc(
      para(undefined, "One") +
        hostPara(undefined, vmlTextBox(para(undefined, "Box a") + para(undefined, "Box b"))),
    );
    const serialized = routingDoc(
      para("0000000A", "One") + hostPara("0000000B", drawingTextBox(para("0000000C", "Box a!"))),
    );

    expect(refusalOf(original, serialized, "0000000C")).toBe(
      "unaligned-paragraph-ordinals: 0000000C",
    );
  });

  test("refuses a paragraph inside mc:AlternateContent, whose Fallback would go stale", () => {
    const original = routingDoc(
      hostPara(
        "B0000001",
        alternateContent(
          drawingTextBox(para("B0000002", "Box")),
          vmlTextBox(para("B0000002", "Box")),
        ),
      ),
    );
    const serialized = routingDoc(hostPara("B0000001", drawingTextBox(para("B0000002", "Box!"))));

    expect(refusalOf(original, serialized, "B0000002")).toBe(
      "paragraph-in-alternate-content: B0000002",
    );
  });

  test("refuses an id the source writes only inside mc:Fallback", () => {
    const original = routingDoc(
      hostPara("B0000001", alternateContent("<w:drawing/>", vmlTextBox(para("F0000001", "Old")))),
    );
    const serialized = routingDoc(hostPara("B0000001", vmlTextBox(para("F0000001", "New"))));

    expect(refusalOf(original, serialized, "F0000001")).toBe("paraId-only-in-fallback: F0000001");
  });

  test("refuses a paragraph whose container changed", () => {
    const original = routingDoc(tableCell(para("C0000001", "Cell")) + para("A0000002", "Tail"));
    const serialized = routingDoc(para("C0000001", "Cell!") + para("A0000002", "Tail"));

    expect(refusalOf(original, serialized, "C0000001")).toBe("container-changed: C0000001");
  });

  test("refuses an authored paragraph that moved within its story", () => {
    const original = routingDoc(
      para("A0000001", "a") + para("A0000002", "b") + para("A0000003", "c"),
    );
    const serialized = routingDoc(
      para("A0000002", "b") + para("A0000001", "a!") + para("A0000003", "c"),
    );

    expect(refusalOf(original, serialized, "A0000001")).toBe("paragraph-moved: A0000001");
  });

  test("refuses a paraId the serialization writes twice", () => {
    const original = routingDoc(para("A0000001", "a") + para("A0000002", "b"));
    const serialized = routingDoc(para("A0000001", "a!") + para("A0000001", "b"));

    expect(refusalOf(original, serialized, "A0000001")).toBe(
      "duplicate-paraId-in-serialized: A0000001",
    );
  });

  test("refuses a part whose WordprocessingML prefix is not w", () => {
    const original =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><x:document xmlns:x="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><x:body><x:p><x:r><x:t>a</x:t></x:r></x:p></x:body></x:document>';
    const serialized = routingDoc(para("0000000A", "a!"));

    expect(refusalOf(original, serialized, "0000000A")).toBe(
      "non-canonical-wordprocessingml-prefix: x:document",
    );
  });

  test("refuses a changed paragraph the serialization leaves unterminated", () => {
    const original = routingDoc(para("A0000001", "a"));
    const serialized = routingDoc(para("A0000001", "a!").replace("</w:p>", ""));

    expect(refusalOf(original, serialized, "A0000001")).toBe("unterminated-paragraph: A0000001");
  });
});

describe("buildPatchedNotePartXml", () => {
  test("binds every serializer namespace when patching an alternate-prefix source", () => {
    const wordNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const word2010Namespace = "http://schemas.microsoft.com/office/word/2010/wordml";
    const word2012Namespace = "http://schemas.microsoft.com/office/word/2012/wordml";
    const compatibilityNamespace = "http://schemas.openxmlformats.org/markup-compatibility/2006";
    const originalXml = `<alt:footnotes xmlns:alt="${wordNamespace}" xmlns:p14="${word2010Namespace}" xmlns:p15="${word2012Namespace}" xmlns:compat="${compatibilityNamespace}"><alt:footnote alt:id="1"><alt:p p14:paraId="P1000001"><alt:r><alt:t>Old</alt:t></alt:r></alt:p></alt:footnote></alt:footnotes>`;
    const baselineXml = `<w:footnotes xmlns:w="${wordNamespace}" xmlns:w14="${word2010Namespace}"><w:footnote w:id="1"><w:p w14:paraId="P1000001"><w:r><w:t>Old</w:t></w:r></w:p></w:footnote></w:footnotes>`;
    const serializedXml = `<w:footnotes xmlns:w="${wordNamespace}" xmlns:w14="${word2010Namespace}"><w:footnote w:id="1"><w:p w14:paraId="P1000001"><w:r><w:t>New</w:t></w:r></w:p></w:footnote></w:footnotes>`;
    const replacementXml = `<w:footnotes xmlns:w="${wordNamespace}" xmlns:w14="${word2010Namespace}" xmlns:w15="${word2012Namespace}" xmlns:mc="${compatibilityNamespace}"><w:footnote w:id="1"><w:p w14:paraId="P1000001"><mc:AlternateContent><mc:Choice Requires="w15"><w:r w15:collapsed="1"><w:t>New</w:t></w:r></mc:Choice></mc:AlternateContent></w:p></w:footnote></w:footnotes>`;

    const patched = patchedXmlOf(
      buildPatchedNotePartXml({
        originalXml,
        baselineXml,
        serializedXml,
        replacementXml,
        elementName: "footnote",
        changedParaIds: new Set(["P1000001"]),
      }),
    );

    expect(patched).toContain("<alt:p");
    expect(patched).toContain(`xmlns:w14="${word2010Namespace}"`);
    expect(patched).toContain(`xmlns:w15="${word2012Namespace}"`);
    expect(patched).toContain(`xmlns:mc="${compatibilityNamespace}"`);
    expect(patched).toContain('<mc:Choice Requires="w15">');
    expect(patched).toContain('<alt:r w15:collapsed="1">');
  });

  // A comment can be anchored on a note's own text, so its range spans the
  // note's paragraphs and an edit inside the span moves a half between them.
  // Replacing only the edited paragraph writes the other half alone.
  test("rewrites the whole note when splicing one paragraph would orphan a comment range", () => {
    const roots =
      'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"';
    const note = (first: string, second: string) =>
      `<w:footnotes ${roots}><w:footnote w:id="1"><w:p w14:paraId="P1000001">${first}</w:p><w:p w14:paraId="P1000002">${second}</w:p></w:footnote></w:footnotes>`;
    // The range opens at the end of the first paragraph, so the only commented
    // text is in the second: the model puts the start there instead.
    const originalXml = note(
      '<w:r><w:t>Opening.</w:t></w:r><w:commentRangeStart w:id="7"/>',
      '<w:r><w:t>Commented.</w:t></w:r><w:commentRangeEnd w:id="7"/>',
    );
    const movedStart = note(
      "<w:r><w:t>Superseded.</w:t></w:r>",
      '<w:commentRangeStart w:id="7"/><w:r><w:t>Commented.</w:t></w:r><w:commentRangeEnd w:id="7"/>',
    );

    const patched = patchedXmlOf(
      buildPatchedNotePartXml({
        originalXml,
        baselineXml: originalXml,
        serializedXml: movedStart,
        replacementXml: movedStart,
        elementName: "footnote",
        changedParaIds: new Set(["P1000001"]),
      }),
    );

    expect(patched).toContain("Superseded.");
    expect(patched).toContain('<w:commentRangeStart w:id="7"/>');
    expect(patched).toContain('<w:commentRangeEnd w:id="7"/>');
  });

  test("refuses a part whose model already lost a comment range half", () => {
    const roots =
      'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"';
    const note = (first: string, second: string) =>
      `<w:footnotes ${roots}><w:footnote w:id="1"><w:p w14:paraId="P1000001">${first}</w:p><w:p w14:paraId="P1000002">${second}</w:p></w:footnote></w:footnotes>`;
    const originalXml = note(
      '<w:commentRangeStart w:id="7"/><w:r><w:t>Opening.</w:t></w:r>',
      '<w:r><w:t>Commented.</w:t></w:r><w:commentRangeEnd w:id="7"/>',
    );
    const withoutStart = note(
      "<w:r><w:t>Superseded.</w:t></w:r>",
      '<w:r><w:t>Commented.</w:t></w:r><w:commentRangeEnd w:id="7"/>',
    );

    expect(
      buildPatchedNotePartXml({
        originalXml,
        baselineXml: originalXml,
        serializedXml: withoutStart,
        replacementXml: withoutStart,
        elementName: "footnote",
        changedParaIds: new Set(["P1000001"]),
      }),
    ).toEqual({ type: "refused", reason: "comment-range-balance" });
  });
});
