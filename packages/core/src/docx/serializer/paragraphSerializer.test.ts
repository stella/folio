import { describe, expect, test } from "bun:test";

import { parseParagraph } from "../paragraphParser";
import type { ComplexField, Paragraph } from "../../types/document";
import { parseXmlDocument, type XmlElement } from "../xmlParser";
import { serializeParagraph } from "./paragraphSerializer";

// The structural runs of a complex field (begin/separate/end) reuse
// `field.formatting`, which the parser captures from the first non-empty
// field-run rPr. `serializeComplexField` selects it with
// `field.formatting ?? field.fieldResult[0]?.formatting`. That `??` is correct
// (not a `keys > 0` guard) because the parser normalizes an empty `<w:rPr/>` to
// `undefined` (runParser.ts `parseRunProperties`), so `field.formatting` is
// never an empty object — there is no reachable case where it is `{}` and the
// result run's formatting should be preferred instead. These round-trips lock
// that: a formatted result run's rPr survives even when the begin run's rPr is
// empty or absent, and a field with no formatting stays `undefined`.
describe("serializeComplexField structural-run formatting round-trip", () => {
  const W_NS =
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"';

  const parseInner = (inner: string): Paragraph => {
    const node = parseXmlDocument(`<w:p ${W_NS}>${inner}</w:p>`) as XmlElement | null;
    if (!node) {
      throw new Error("failed to parse paragraph fixture");
    }
    return parseParagraph(node, null, null, null, null, null);
  };

  // `serializeParagraph` emits `<w:p>` without namespace declarations; inject
  // them so the serialized paragraph can be parsed back for comparison.
  const reparse = (paragraphXml: string): Paragraph => {
    const node = parseXmlDocument(
      paragraphXml.replace(/^<w:p(?=[\s>])/u, `<w:p ${W_NS}`),
    ) as XmlElement | null;
    if (!node) {
      throw new Error("failed to reparse serialized paragraph");
    }
    return parseParagraph(node, null, null, null, null, null);
  };

  const fieldOf = (paragraph: Paragraph): ComplexField => {
    const field = paragraph.content.find((c): c is ComplexField => c.type === "complexField");
    if (!field) {
      throw new Error("no complex field in paragraph");
    }
    return field;
  };

  test("separate-run rPr distinct from the result run round-trips (the fixed bug)", () => {
    // Mirrors the podily REF field: the begin run's rPr is empty, the separate
    // run carries the field's font (Georgia), and the visible result run has a
    // different rPr (blue). `field.formatting` must come from the separate run,
    // and serializing structural runs with it (not the result formatting) is
    // what makes Georgia survive. Emitting the result formatting instead — the
    // pre-fix behavior — would recapture blue and drop Georgia.
    const original = parseInner(`
      <w:r><w:rPr/><w:fldChar w:fldCharType="begin"/></w:r>
      <w:r><w:instrText xml:space="preserve"> REF _Ref1 \\h </w:instrText></w:r>
      <w:r><w:rPr><w:rFonts w:ascii="Georgia" w:hAnsi="Georgia"/><w:sz w:val="21"/></w:rPr><w:fldChar w:fldCharType="separate"/></w:r>
      <w:r><w:rPr><w:color w:val="0000FF"/></w:rPr><w:t>5.2</w:t></w:r>
      <w:r><w:fldChar w:fldCharType="end"/></w:r>
    `);
    const field = fieldOf(original);
    expect(field.formatting?.fontFamily).toEqual({ ascii: "Georgia", hAnsi: "Georgia" });
    // The guard only bites when the captured field formatting differs from the
    // result run's formatting.
    expect(field.formatting).not.toEqual(field.fieldResult[0]?.formatting);

    const roundTripped = fieldOf(reparse(serializeParagraph(original)));
    expect(roundTripped.formatting).toEqual(field.formatting);
  });

  test("begin run with no rPr and a formatted result run round-trips", () => {
    const original = parseInner(`
      <w:r><w:fldChar w:fldCharType="begin"/></w:r>
      <w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>
      <w:r><w:fldChar w:fldCharType="separate"/></w:r>
      <w:r><w:rPr><w:color w:val="00FF00"/></w:rPr><w:t>1</w:t></w:r>
      <w:r><w:fldChar w:fldCharType="end"/></w:r>
    `);
    const field = fieldOf(original);
    expect(field.formatting?.color).toEqual({ rgb: "00FF00" });

    const roundTripped = fieldOf(reparse(serializeParagraph(original)));
    expect(roundTripped.formatting).toEqual(field.formatting);
  });

  test("collapsed PAGE field (no result run) round-trips its rPr (#909)", () => {
    const original = parseInner(`
      <w:r>
        <w:rPr><w:color w:val="595959"/><w:sz w:val="14"/></w:rPr>
        <w:fldChar w:fldCharType="begin"/>
        <w:instrText xml:space="preserve"> PAGE </w:instrText>
        <w:fldChar w:fldCharType="separate"/>
        <w:fldChar w:fldCharType="end"/>
      </w:r>
    `);
    const field = fieldOf(original);
    expect(field.fieldResult).toHaveLength(0);
    expect(field.formatting).toEqual({ color: { rgb: "595959" }, fontSize: 14 });

    const roundTripped = fieldOf(reparse(serializeParagraph(original)));
    expect(roundTripped.formatting).toEqual(field.formatting);
  });

  test("a field with no run formatting keeps field.formatting undefined (never {})", () => {
    const original = parseInner(`
      <w:r><w:rPr/><w:fldChar w:fldCharType="begin"/></w:r>
      <w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>
      <w:r><w:fldChar w:fldCharType="separate"/></w:r>
      <w:r><w:t>1</w:t></w:r>
      <w:r><w:fldChar w:fldCharType="end"/></w:r>
    `);
    const field = fieldOf(original);
    // The empty begin rPr normalizes to undefined, not {}, so the `??` fallback
    // never has to choose between an empty object and the result formatting.
    expect(field.formatting).toBeUndefined();

    const roundTripped = fieldOf(reparse(serializeParagraph(original)));
    expect(roundTripped.formatting).toBeUndefined();
  });
});

describe("serializeParagraph tracked-change hardening", () => {
  test("serializes deletion runs using delText and delInstrText", () => {
    const paragraph: Paragraph = {
      type: "paragraph",
      content: [
        {
          type: "deletion",
          info: { id: 11, author: "Reviewer", date: "2026-02-22T10:00:00Z" },
          content: [
            {
              type: "run",
              content: [
                { type: "text", text: "Removed" },
                { type: "instrText", text: " MERGEFIELD name " },
              ],
            },
          ],
        },
      ],
    };

    const xml = serializeParagraph(paragraph);

    expect(xml).toContain('<w:del w:id="11" w:author="Reviewer" w:date="2026-02-22T10:00:00Z">');
    expect(xml).toContain("<w:delText>Removed</w:delText>");
    expect(xml).toContain(
      '<w:delInstrText xml:space="preserve"> MERGEFIELD name </w:delInstrText>',
    );
    expect(xml).not.toContain("<w:t>Removed</w:t>");
    expect(xml).not.toContain('<w:instrText xml:space="preserve"> MERGEFIELD name </w:instrText>');
  });

  test("normalizes invalid tracked-change metadata while serializing", () => {
    const paragraph: Paragraph = {
      type: "paragraph",
      content: [
        {
          type: "insertion",
          info: { id: -5, author: "   ", date: "   " },
          content: [
            {
              type: "run",
              content: [{ type: "text", text: "Added" }],
            },
          ],
        },
      ],
    };

    const xml = serializeParagraph(paragraph);

    expect(xml).toContain('<w:ins w:id="0" w:author="Unknown">');
    expect(xml).not.toContain("w:date=");
  });

  test("folds an out-of-range revision id into the signed 32-bit range (eigenpal #1093)", () => {
    // A `Date.now()`-derived id (~1.8e12) is a well-formed positive integer, so
    // the invalid/negative guard used to pass it straight through to `w:id`.
    const paragraph: Paragraph = {
      type: "paragraph",
      content: [
        {
          type: "insertion",
          info: { id: 1_784_212_345_678, author: "Reviewer" },
          content: [{ type: "run", content: [{ type: "text", text: "Added" }] }],
        },
      ],
    };

    const xml = serializeParagraph(paragraph);
    const id = Number(/<w:ins w:id="(\d+)"/u.exec(xml)?.[1]);
    expect(id).toBeLessThanOrEqual(2_147_483_647);
    expect(id).toBeGreaterThanOrEqual(0);
  });

  test("keeps distinct out-of-range ids distinct so revisions do not merge", () => {
    const idFor = (id: number): string => {
      const paragraph: Paragraph = {
        type: "paragraph",
        content: [
          {
            type: "insertion",
            info: { id, author: "Reviewer" },
            content: [{ type: "run", content: [{ type: "text", text: "Added" }] }],
          },
        ],
      };
      return /<w:ins w:id="(\d+)"/u.exec(serializeParagraph(paragraph))?.[1] ?? "";
    };

    expect(idFor(1_784_212_345_678)).not.toBe(idFor(1_784_212_345_679));
  });
});

describe("serializeParagraph native frame geometry", () => {
  test("writes frame wrap spacing alongside its size and anchors", () => {
    const paragraph: Paragraph = {
      type: "paragraph",
      formatting: {
        frame: {
          dropCap: "none",
          lines: 2,
          width: 3600,
          height: 1440,
          hSpace: 144,
          vSpace: 72,
          hAnchor: "page",
          vAnchor: "text",
          x: 720,
          y: 144,
          wrap: "around",
        },
      },
      content: [{ type: "run", content: [{ type: "text", text: "Framed text" }] }],
    };

    const xml = serializeParagraph(paragraph);

    expect(xml).toContain(
      '<w:framePr w:dropCap="none" w:lines="2" w:w="3600" w:h="1440" w:hSpace="144" w:vSpace="72" w:hAnchor="page" w:vAnchor="text" w:x="720" w:y="144" w:wrap="around"/>',
    );
  });
});
