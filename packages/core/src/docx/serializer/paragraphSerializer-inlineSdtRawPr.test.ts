/**
 * Round-trip guard for inline SDT (`w:sdt`) property fidelity.
 *
 * The inline serializer used to re-synthesize `<w:sdtPr>` from the modeled
 * projection alone, silently dropping every unmodeled OOXML feature (`w:id`,
 * `w:dataBinding`, `w15:*`, custom XML mappings) and `<w:sdtEndPr>` on save.
 * These tests lock the fix: the captured raw properties are replayed verbatim
 * (mirroring the block-SDT serializer), while a modeled interactive edit is
 * still reconciled into the raw string before replay.
 */

import { describe, expect, test } from "bun:test";

import { fromProseDoc } from "../../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../../prosemirror/conversion/toProseDoc";
import { schema } from "../../prosemirror/schema";
import type { Document, InlineSdt, Paragraph } from "../../types/document";
import { parseParagraph } from "../paragraphParser";
import type { XmlElement } from "../xmlParser";
import { parseXmlDocument } from "../xmlParser";
import { serializeParagraph } from "./paragraphSerializer";

const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const W15_NS = 'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"';

function parseParagraphXml(xml: string): Paragraph {
  const root = parseXmlDocument(xml) as XmlElement | null;
  if (!root) {
    throw new Error("Failed to parse paragraph XML fixture");
  }
  return parseParagraph(root, null, null, null, null, null);
}

function firstInlineSdt(paragraph: Paragraph): InlineSdt {
  const sdt = paragraph.content.find((content) => content.type === "inlineSdt");
  if (sdt?.type !== "inlineSdt") {
    throw new Error("Expected an inlineSdt in paragraph content");
  }
  return sdt;
}

describe("inline SDT raw-property round-trip", () => {
  test("preserves unmodeled w:sdtPr features and w:sdtEndPr through parse → serialize → re-parse", () => {
    const xml =
      `<w:p ${W_NS} ${W15_NS}>` +
      "<w:r><w:t>before </w:t></w:r>" +
      "<w:sdt><w:sdtPr>" +
      '<w:alias w:val="Party Name"/><w:tag w:val="party"/><w:id w:val="123456789"/>' +
      '<w:dataBinding w:xpath="/ns0:root/ns0:party" w:storeItemID="{GUID}"/>' +
      "<w:text/>" +
      "</w:sdtPr>" +
      "<w:sdtEndPr><w:rPr><w:b/></w:rPr></w:sdtEndPr>" +
      "<w:sdtContent><w:r><w:t>Acme Corp</w:t></w:r></w:sdtContent>" +
      "</w:sdt>" +
      "<w:r><w:t> after</w:t></w:r></w:p>";

    const serialized = serializeParagraph(parseParagraphXml(xml));

    // Unmodeled markers and the end-properties block must survive save.
    expect(serialized).toContain('<w:id w:val="123456789"/>');
    expect(serialized).toContain("<w:dataBinding");
    expect(serialized).toContain('w:xpath="/ns0:root/ns0:party"');
    expect(serialized).toContain('w:storeItemID="{GUID}"');
    expect(serialized).toContain("<w:sdtEndPr><w:rPr><w:b/></w:rPr></w:sdtEndPr>");
    expect(serialized).toContain("Acme Corp");

    // Re-parsing the saved XML must recover an equivalent inline SDT: same
    // inner text, same verbatim raw properties, same modeled projection.
    const reparsed = firstInlineSdt(
      parseParagraphXml(
        `<w:p ${W_NS} ${W15_NS}>${serialized.replace(/^<w:p>/u, "").replace(/<\/w:p>$/u, "")}</w:p>`,
      ),
    );
    expect(reparsed.properties.id).toBe(123456789);
    expect(reparsed.properties.alias).toBe("Party Name");
    expect(reparsed.properties.tag).toBe("party");
    expect(reparsed.properties.rawPropertiesXml).toContain("<w:dataBinding");
    expect(reparsed.properties.rawEndPropertiesXml).toContain("<w:rPr><w:b/></w:rPr>");
    const innerText = reparsed.content
      .flatMap((c) => (c.type === "run" ? c.content : []))
      .map((rc) => (rc.type === "text" ? rc.text : ""))
      .join("");
    expect(innerText).toBe("Acme Corp");
  });

  test("reconciles a modeled checkbox toggle into the replayed raw properties", () => {
    const xml =
      `<w:p ${W_NS} xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">` +
      "<w:sdt><w:sdtPr>" +
      '<w:tag w:val="agree"/>' +
      '<w14:checkbox><w14:checked w14:val="0"/></w14:checkbox>' +
      "</w:sdtPr>" +
      "<w:sdtContent><w:r><w:t>☐</w:t></w:r></w:sdtContent>" +
      "</w:sdt></w:p>";

    const paragraph = parseParagraphXml(xml);
    const sdt = firstInlineSdt(paragraph);
    expect(sdt.properties.sdtType).toBe("checkbox");
    expect(sdt.properties.checked).toBe(false);

    // Flip the modeled state (as an interactive checkbox toggle would) and
    // confirm the replayed raw XML reflects the new value, not the stale one.
    sdt.properties.checked = true;
    const serialized = serializeParagraph(paragraph);
    expect(serialized).toContain('w14:val="1"');
    expect(serialized).not.toContain('w14:val="0"');
    // The unmodeled tag inside the raw properties is untouched.
    expect(serialized).toContain('<w:tag w:val="agree"/>');
  });

  test("preserves raw w:sdtPr through the full ProseMirror round-trip", () => {
    const xml =
      `<w:p ${W_NS}>` +
      "<w:sdt><w:sdtPr>" +
      '<w:alias w:val="clause"/><w:id w:val="42"/>' +
      '<w:dataBinding w:xpath="/root/clause" w:storeItemID="{ABC}"/>' +
      "<w:text/>" +
      "</w:sdtPr>" +
      "<w:sdtContent><w:r><w:t>Value</w:t></w:r></w:sdtContent>" +
      "</w:sdt></w:p>";

    const paragraph = parseParagraphXml(xml);
    const baseDocument = {
      package: { document: { content: [paragraph] } } as Document["package"],
    } as Document;

    const roundTripped = fromProseDoc(toProseDoc(baseDocument), baseDocument);
    const rtParagraph = roundTripped.package.document.content.find(
      (c): c is Paragraph => c.type === "paragraph",
    );
    expect(rtParagraph).toBeDefined();
    if (!rtParagraph) {
      return;
    }

    const serialized = serializeParagraph(rtParagraph);
    expect(serialized).toContain('<w:id w:val="42"/>');
    expect(serialized).toContain("<w:dataBinding");
    expect(serialized).toContain('w:xpath="/root/clause"');
    expect(serialized).toContain("Value");
  });
});

describe("inline SDT raw-property structural validation", () => {
  test("falls back to synthesized sdtPr when rawPropertiesXml closes the SDT early", () => {
    // `rawPropertiesXml` is normally a verbatim `<w:sdtPr>` snapshot captured
    // by our own parser, but it can also arrive from an untrusted surface
    // (a programmatically constructed node, a collaboration payload). A
    // value like this one closes `<w:sdtContent>`/`<w:sdt>` early and
    // splices in sibling markup — it must never be spliced into the
    // serialized document verbatim.
    const malicious =
      '<w:sdtPr><w:tag w:val="x"/></w:sdtPr></w:sdtContent></w:sdt>' +
      "<w:p><w:r><w:t>INJECTED</w:t></w:r></w:p>" +
      "<w:sdt><w:sdtPr>";

    const paragraph: Paragraph = {
      type: "paragraph",
      content: [
        {
          type: "inlineSdt",
          properties: { sdtType: "richText", alias: "clause", rawPropertiesXml: malicious },
          content: [{ type: "run", content: [{ type: "text", text: "value" }] }],
        },
      ],
    };

    const serialized = serializeParagraph(paragraph);

    expect(serialized).not.toContain("INJECTED");
    expect(serialized).not.toContain(malicious);
    // Falls back to a single well-formed synthesized <w:sdtPr>, still
    // carrying the modeled alias.
    expect(serialized).toContain('<w:sdtPr><w:alias w:val="clause"/></w:sdtPr>');
  });

  test("falls back to no end-properties when rawEndPropertiesXml isn't a single well-formed <w:sdtEndPr>", () => {
    const paragraph: Paragraph = {
      type: "paragraph",
      content: [
        {
          type: "inlineSdt",
          properties: {
            sdtType: "richText",
            rawEndPropertiesXml: "<w:sdtEndPr/><w:p><w:r><w:t>INJECTED</w:t></w:r></w:p>",
          },
          content: [{ type: "run", content: [{ type: "text", text: "value" }] }],
        },
      ],
    };

    const serialized = serializeParagraph(paragraph);

    expect(serialized).not.toContain("INJECTED");
    expect(serialized).not.toContain("<w:sdtEndPr");
  });

  test("still replays a well-formed rawPropertiesXml/rawEndPropertiesXml verbatim", () => {
    const paragraph: Paragraph = {
      type: "paragraph",
      content: [
        {
          type: "inlineSdt",
          properties: {
            sdtType: "richText",
            rawPropertiesXml: '<w:sdtPr><w:tag w:val="ok"/></w:sdtPr>',
            rawEndPropertiesXml: "<w:sdtEndPr><w:rPr><w:b/></w:rPr></w:sdtEndPr>",
          },
          content: [{ type: "run", content: [{ type: "text", text: "value" }] }],
        },
      ],
    };

    const serialized = serializeParagraph(paragraph);

    expect(serialized).toContain('<w:sdtPr><w:tag w:val="ok"/></w:sdtPr>');
    expect(serialized).toContain("<w:sdtEndPr><w:rPr><w:b/></w:rPr></w:sdtEndPr>");
  });
});

describe("inline SDT null-default normalization", () => {
  test("PM null-default attrs never enter the model or the serialized XML", () => {
    // A control with no id / dropdownLastValue / checked: `toProseDoc`
    // materializes these as PM null defaults on the `sdt` node.
    const original: Paragraph = {
      type: "paragraph",
      content: [
        {
          type: "inlineSdt",
          properties: { sdtType: "richText", alias: "x" },
          content: [{ type: "run", content: [{ type: "text", text: "v" }] }],
        },
      ],
    };
    const baseDocument = {
      package: { document: { content: [original] } } as Document["package"],
    } as Document;

    // Confirm the scenario: the PM `sdt` node really carries null defaults.
    const pmDoc = toProseDoc(baseDocument);
    let sdtNode: { attrs: Record<string, unknown> } | undefined;
    pmDoc.descendants((node) => {
      if (node.type.name === "sdt") {
        sdtNode = node;
        return false;
      }
      return true;
    });
    expect(sdtNode?.attrs["id"]).toBeNull();
    expect(sdtNode?.attrs["dropdownLastValue"]).toBeNull();
    expect(sdtNode?.attrs["checked"]).toBeNull();

    // The null defaults must not survive back into the typed model.
    const roundTripped = fromProseDoc(pmDoc, baseDocument);
    const rtParagraph = roundTripped.package.document.content.find(
      (c): c is Paragraph => c.type === "paragraph",
    );
    const sdt = rtParagraph ? firstInlineSdt(rtParagraph) : undefined;
    expect(sdt?.properties.id).toBeUndefined();
    expect(sdt?.properties.dropdownLastValue).toBeUndefined();
    expect(sdt?.properties.checked).toBeUndefined();
    expect(sdt?.properties.id).not.toBeNull();

    // And no stray `null` reaches the serialized XML (e.g. `<w:id w:val="null"/>`).
    const serialized = rtParagraph ? serializeParagraph(rtParagraph) : "";
    expect(serialized).not.toContain("<w:id");
    expect(serialized).not.toContain("null");
  });

  test("toDOM omits data-sdt-id when the id is a null default", () => {
    const toDOM = schema.nodes["sdt"]?.spec.toDOM;
    expect(toDOM).toBeDefined();
    if (!toDOM) {
      return;
    }
    const node = schema.node("sdt", {}, undefined);
    const rendered = toDOM(node) as [string, Record<string, string>, number];
    const domAttrs = rendered[1];
    expect("data-sdt-id" in domAttrs).toBe(false);
    expect(Object.values(domAttrs)).not.toContain("null");
  });
});
