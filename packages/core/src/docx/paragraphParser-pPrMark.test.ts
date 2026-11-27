import { describe, expect, test } from "bun:test";

import { getTrackedChangesFromDoc } from "../ai-edits/read";
import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import { parseParagraph } from "./paragraphParser";
import { FOLIO_REVIEW_HISTORY_NAMESPACE } from "./reviewHistoryNamespace";
import { serializeParagraph } from "./serializer/paragraphSerializer";
import type { XmlElement } from "./xmlParser";
import { parseXmlDocument } from "./xmlParser";

function parseParagraphXml(xml: string) {
  const root = parseXmlDocument(xml) as XmlElement | null;
  if (!root) {
    throw new Error("Failed to parse paragraph XML fixture");
  }
  return parseParagraph(root, null, null, null, null, null);
}

describe("parseParagraph — paragraph-mark tracked change (ECMA-376 §17.13.5)", () => {
  const namespace =
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
  test("reads and writes a paragraph mark run-property change once after the live properties", () => {
    const source = `<w:p ${namespace}><w:pPr><w:rPr><w:ins w:id="1" w:author="A"/><w:i/><w:rPrChange w:id="2" w:author="B"><w:rPr><w:b/><w:specVanish/></w:rPr></w:rPrChange></w:rPr></w:pPr></w:p>`;
    const paragraph = parseParagraphXml(source);
    expect(paragraph.markPropertyChange).toStrictEqual({
      info: { id: 2, author: "B" },
      previousFormatting: {
        runProperties: { bold: true },
        runInWithNext: true,
      },
    });
    expect(paragraph.formatting).toStrictEqual({
      runProperties: { italic: true },
    });
    const xml = serializeParagraph(paragraph);
    expect(xml.match(/<w:rPrChange\b/gu)).toHaveLength(1);
    expect(xml).toContain('<w:i/><w:rPrChange w:id="2"');
    const reopened = parseParagraphXml(
      xml.replace("<w:p", `<w:p ${namespace}`),
    );
    expect(reopened.markPropertyChange).toStrictEqual(
      paragraph.markPropertyChange,
    );
    expect(reopened.pPrMark).toStrictEqual(paragraph.pPrMark);

    const document = { package: { document: { content: [paragraph] } } };
    const view = toProseDoc(document);
    expect(getTrackedChangesFromDoc(view)).toContainEqual(
      expect.objectContaining({ id: 2, type: "paragraphPropertiesChanged" }),
    );
    const rebuilt = fromProseDoc(view, document, {
      stylesheetSource: { type: "package" },
    }).package.document.content.at(0);
    expect(
      rebuilt?.type === "paragraph" && rebuilt.markPropertyChange,
    ).toStrictEqual(paragraph.markPropertyChange);
  });

  test("reopens a written mark property revision whose original formatting was absent", () => {
    const change = { info: { id: 2, author: "Reviewer" } };
    const xml = serializeParagraph({
      type: "paragraph",
      content: [],
      formatting: { runInWithNext: true },
      markPropertyChange: change,
    });
    const history = `xmlns:folio="${FOLIO_REVIEW_HISTORY_NAMESPACE}"`;
    const reopened = parseParagraphXml(
      xml.replace("<w:p", `<w:p ${namespace} ${history}`),
    );
    expect(reopened.markPropertyChange).toStrictEqual(change);
  });

  test("reads <w:pPr><w:rPr><w:ins/> as pPrMark.kind = 'ins'", () => {
    const paragraph = parseParagraphXml(`
      <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:pPr>
          <w:rPr>
            <w:ins w:id="42" w:author="Alice" w:date="2026-05-01T10:00:00Z"/>
          </w:rPr>
        </w:pPr>
      </w:p>
    `);

    expect(paragraph.pPrMark).toEqual({
      kind: "ins",
      info: { id: 42, author: "Alice", date: "2026-05-01T10:00:00Z" },
    });
  });

  test("reads <w:pPr><w:rPr><w:del/> as pPrMark.kind = 'del'", () => {
    const paragraph = parseParagraphXml(`
      <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:pPr>
          <w:rPr>
            <w:del w:id="7" w:author="Bob" w:date="2026-05-02T11:00:00Z"/>
          </w:rPr>
        </w:pPr>
      </w:p>
    `);

    expect(paragraph.pPrMark).toEqual({
      kind: "del",
      info: { id: 7, author: "Bob", date: "2026-05-02T11:00:00Z" },
    });
  });

  test("leaves pPrMark unset when neither ins nor del is present in rPr", () => {
    const paragraph = parseParagraphXml(`
      <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:pPr>
          <w:rPr><w:b/></w:rPr>
        </w:pPr>
      </w:p>
    `);

    expect(paragraph.pPrMark).toBeUndefined();
  });

  test("reads an inserted mark's later deletion beside its insertion", () => {
    const paragraph = parseParagraphXml(`
      <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:pPr>
          <w:rPr>
            <w:ins w:id="1" w:author="A"/>
            <w:del w:id="2" w:author="B"/>
          </w:rPr>
        </w:pPr>
      </w:p>
    `);

    expect(paragraph.pPrMark).toEqual({
      kind: "ins",
      info: { id: 1, author: "A" },
      deletion: { info: { id: 2, author: "B" } },
    });
  });

  test("reads a moved-to mark's later deletion beside it", () => {
    const paragraph = parseParagraphXml(`
      <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:pPr>
          <w:rPr>
            <w:del w:id="4" w:author="B"/>
            <w:moveTo w:id="3" w:author="A"/>
          </w:rPr>
        </w:pPr>
      </w:p>
    `);

    expect(paragraph.pPrMark).toEqual({
      kind: "moveTo",
      info: { id: 3, author: "A" },
      deletion: { info: { id: 4, author: "B" } },
    });
  });

  test("omits w:date when it is absent on the source element", () => {
    const paragraph = parseParagraphXml(`
      <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:pPr>
          <w:rPr>
            <w:ins w:id="3" w:author="Carol"/>
          </w:rPr>
        </w:pPr>
      </w:p>
    `);

    expect(paragraph.pPrMark).toEqual({
      kind: "ins",
      info: { id: 3, author: "Carol" },
    });
  });

  test("reads paragraph-mark changes in the Strict WordprocessingML namespace", () => {
    const paragraph = parseParagraphXml(`
      <x:p xmlns:x="http://purl.oclc.org/ooxml/wordprocessingml/main">
        <x:pPr>
          <x:rPr>
            <x:moveTo x:id="4" x:author="Dana"/>
          </x:rPr>
        </x:pPr>
      </x:p>
    `);

    expect(paragraph.pPrMark).toEqual({
      kind: "moveTo",
      info: { id: 4, author: "Dana" },
    });
  });

  test("ignores same-named paragraph-mark elements from a foreign namespace", () => {
    const paragraph = parseParagraphXml(`
      <w:p
        xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        xmlns:foreign="https://example.com/foreign"
      >
        <w:pPr>
          <w:rPr>
            <foreign:moveFrom w:id="5" w:author="Eve"/>
            <foreign:del w:id="6" w:author="Frank"/>
          </w:rPr>
        </w:pPr>
      </w:p>
    `);

    expect(paragraph.pPrMark).toBeUndefined();
  });
});
