import { describe, expect, test } from "bun:test";
import { panic } from "better-result";

import { parseParagraph } from "./paragraphParser";
import { parseSectionProperties } from "./sectionParser";
import { serializeParagraph } from "./serializer/paragraphSerializer";
import { serializeSectionProperties } from "./serializer/sectionPropertiesSerializer";
import { serializeTable } from "./serializer/tableSerializer";
import { serializeTrackedChangeAttributes } from "./serializer/trackedChangeAttributes";
import { parseTable } from "./tableParser";
import { DATE_UTC_NAMESPACE_URI, parseTrackedChangeInfo } from "./trackedChangeInfo";
import { parseXmlDocument, type XmlElement } from "./xmlParser";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const NS = `xmlns:w="${W}" xmlns:du="${DATE_UTC_NAMESPACE_URI}"`;
const UTC_DATE = "2026-09-08T08:07:06Z";

const parseElement = (xml: string): XmlElement => {
  const element = parseXmlDocument(xml);
  if (!element) {
    panic("expected XML element");
  }
  return element;
};

const replaceUtcDatePrefix = (
  ...infos: Array<{ utcDate?: { attribute: string; value: string } } | undefined>
): void => {
  for (const info of infos) {
    if (info?.utcDate) {
      info.utcDate.attribute = "unbound:dateUtc";
    }
  }
};

describe("tracked-change metadata round-trip", () => {
  test("preserves a remapped dateUtc attribute across paragraph and run revision sites", () => {
    const paragraph = parseParagraph(
      parseElement(`<w:p ${NS}>
        <w:pPr>
          <w:jc w:val="center"/>
          <w:rPr><w:moveTo w:id="1" w:author="Reviewer" du:dateUtc="${UTC_DATE}"/></w:rPr>
          <w:pPrChange w:id="2" w:author="Reviewer" du:dateUtc="${UTC_DATE}">
            <w:pPr><w:jc w:val="left"/></w:pPr>
          </w:pPrChange>
        </w:pPr>
        <w:r>
          <w:rPr>
            <w:i/>
            <w:rPrChange w:id="3" w:author="Reviewer" du:dateUtc="${UTC_DATE}">
              <w:rPr><w:b/></w:rPr>
            </w:rPrChange>
          </w:rPr>
          <w:t>current</w:t>
        </w:r>
        <w:ins w:id="4" w:author="Reviewer" du:dateUtc="${UTC_DATE}">
          <w:r><w:t>inserted</w:t></w:r>
        </w:ins>
      </w:p>`),
      null,
      null,
      null,
      null,
      null,
    );

    expect(paragraph.pPrMark?.info.utcDate).toEqual({
      attribute: "w16du:dateUtc",
      value: UTC_DATE,
    });
    expect(paragraph.propertyChanges?.at(0)?.info.utcDate?.attribute).toBe("w16du:dateUtc");
    const currentRun = paragraph.content.find((content) => content.type === "run");
    expect(
      currentRun?.type === "run" ? currentRun.propertyChanges?.at(0)?.info.utcDate : null,
    ).toEqual({ attribute: "w16du:dateUtc", value: UTC_DATE });
    const insertion = paragraph.content.find((content) => content.type === "insertion");
    expect(insertion?.type === "insertion" ? insertion.info.utcDate : null).toEqual({
      attribute: "w16du:dateUtc",
      value: UTC_DATE,
    });

    replaceUtcDatePrefix(
      paragraph.pPrMark?.info,
      paragraph.propertyChanges?.at(0)?.info,
      currentRun?.type === "run" ? currentRun.propertyChanges?.at(0)?.info : undefined,
      insertion?.type === "insertion" ? insertion.info : undefined,
    );

    const serialized = serializeParagraph(paragraph);
    expect(serialized.match(/w16du:dateUtc=/gu)).toHaveLength(4);
    expect(serialized).not.toMatch(/\sdu:dateUtc=/u);
    expect(serialized).not.toContain("unbound:dateUtc");
  });

  test("preserves a remapped dateUtc attribute on section property changes", () => {
    const section = parseSectionProperties(
      parseElement(`<w:sectPr ${NS}>
        <w:pgSz w:w="11906" w:h="16838"/>
        <w:sectPrChange w:id="5" w:author="Reviewer" du:dateUtc="${UTC_DATE}">
          <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
        </w:sectPrChange>
      </w:sectPr>`),
    );

    expect(section.propertyChanges?.at(0)?.info.utcDate).toEqual({
      attribute: "w16du:dateUtc",
      value: UTC_DATE,
    });
    replaceUtcDatePrefix(section.propertyChanges?.at(0)?.info);
    const serialized = serializeSectionProperties(section);
    expect(serialized).toContain(`w16du:dateUtc="${UTC_DATE}"`);
    expect(serialized).not.toContain("unbound:dateUtc");
  });

  test("preserves a remapped dateUtc attribute across table revision sites", () => {
    const table = parseTable(
      parseElement(`<w:tbl ${NS}>
        <w:tblPr>
          <w:tblW w:w="5000" w:type="dxa"/>
          <w:tblPrChange w:id="6" w:author="Reviewer" du:dateUtc="${UTC_DATE}">
            <w:tblPr><w:tblW w:w="4000" w:type="dxa"/></w:tblPr>
          </w:tblPrChange>
        </w:tblPr>
        <w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>
        <w:tr>
          <w:trPr>
            <w:trHeight w:val="300"/>
            <w:ins w:id="7" w:author="Reviewer" du:dateUtc="${UTC_DATE}"/>
            <w:trPrChange w:id="8" w:author="Reviewer" du:dateUtc="${UTC_DATE}">
              <w:trPr><w:trHeight w:val="200"/></w:trPr>
            </w:trPrChange>
          </w:trPr>
          <w:tc>
            <w:tcPr>
              <w:tcW w:w="5000" w:type="dxa"/>
              <w:cellIns w:id="9" w:author="Reviewer" du:dateUtc="${UTC_DATE}"/>
              <w:tcPrChange w:id="10" w:author="Reviewer" du:dateUtc="${UTC_DATE}">
                <w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>
              </w:tcPrChange>
            </w:tcPr>
            <w:p/>
          </w:tc>
        </w:tr>
      </w:tbl>`),
      null,
      null,
      null,
      null,
      new Map(),
    );
    if (!table) {
      panic("expected table");
    }

    expect(table.propertyChanges?.at(0)?.info.utcDate?.attribute).toBe("w16du:dateUtc");
    expect(table.rows.at(0)?.propertyChanges?.at(0)?.info.utcDate?.attribute).toBe("w16du:dateUtc");
    expect(table.rows.at(0)?.structuralChange?.info.utcDate?.attribute).toBe("w16du:dateUtc");
    expect(table.rows.at(0)?.cells.at(0)?.propertyChanges?.at(0)?.info.utcDate?.attribute).toBe(
      "w16du:dateUtc",
    );
    expect(table.rows.at(0)?.cells.at(0)?.structuralChange?.info.utcDate?.attribute).toBe(
      "w16du:dateUtc",
    );

    replaceUtcDatePrefix(
      table.propertyChanges?.at(0)?.info,
      table.rows.at(0)?.propertyChanges?.at(0)?.info,
      table.rows.at(0)?.structuralChange?.info,
      table.rows.at(0)?.cells.at(0)?.propertyChanges?.at(0)?.info,
      table.rows.at(0)?.cells.at(0)?.structuralChange?.info,
    );

    const serialized = serializeTable(table, serializeParagraph);
    expect(serialized.match(/w16du:dateUtc=/gu)).toHaveLength(5);
    expect(serialized).not.toMatch(/\sdu:dateUtc=/u);
    expect(serialized).not.toContain("unbound:dateUtc");
  });

  test("does not reinterpret a foreign same-local-name attribute as dateUtc metadata", () => {
    const info = parseTrackedChangeInfo(
      parseElement(
        `<w:ins xmlns:w="${W}" xmlns:x="urn:example:foreign" w:id="11" w:author="Reviewer" x:dateUtc="${UTC_DATE}"/>`,
      ),
    );

    expect(info.utcDate).toBeUndefined();
  });

  test("ignores malformed UTC metadata and never trusts its lexical attribute name", () => {
    expect(
      serializeTrackedChangeAttributes({
        id: 12,
        author: "Reviewer",
        utcDate: { attribute: "unbound:dateUtc", value: UTC_DATE },
      }),
    ).toContain(`w16du:dateUtc="${UTC_DATE}"`);
    const malformed = {
      id: 13,
      author: "Reviewer",
      utcDate: { attribute: "unbound:dateUtc", value: UTC_DATE },
    };
    Reflect.set(malformed.utcDate, "value", 42);
    expect(serializeTrackedChangeAttributes(malformed)).toBe('w:id="13" w:author="Reviewer"');
  });
});
