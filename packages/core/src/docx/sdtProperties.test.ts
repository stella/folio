/**
 * Tests for the shared SDT-properties parser.
 *
 * Focus areas not covered by the higher-level block-SDT parser tests:
 * recognising `w14:` / `w15:` prefixed marker elements (the most visible
 * one being `w14:checkbox`).
 */

import { describe, expect, test } from "bun:test";

import { parseSdtProperties } from "./sdtProperties";
import { serializeSdtProperties } from "./serializer/sdtPropertiesSerializer";
import { parseXml } from "./xmlParser";

/** The markup folio writes back for a parsed property set. */
function rewrite(xml: string): string {
  return serializeSdtProperties(parseSdtPrXml(xml));
}

function parseSdtPrXml(xml: string) {
  const root = parseXml(xml);
  const sdtPr = root.elements?.[0];
  if (!sdtPr) {
    throw new TypeError("expected sdtPr root");
  }
  return parseSdtProperties(sdtPr);
}

describe("parseSdtProperties — prefixed marker elements", () => {
  test("recognizes w14:checkbox and returns sdtType=checkbox", () => {
    const props = parseSdtPrXml(
      '<w:sdtPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:tag w:val="agree"/><w14:checkbox><w14:checked w14:val="1"/></w14:checkbox></w:sdtPr>',
    );
    expect(props.sdtType).toBe("checkbox");
    expect(props.checked).toBe(true);
    expect(props.tag).toBe("agree");
  });

  test("recognizes a w:date marker even on prefixed children", () => {
    const props = parseSdtPrXml(
      '<w:sdtPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:date w:fullDate="2026-06-02T00:00:00Z"><w:dateFormat w:val="d MMMM yyyy"/></w:date></w:sdtPr>',
    );
    expect(props.sdtType).toBe("date");
    expect(props.dateFormat).toBe("d MMMM yyyy");
    expect(props.dateValueISO).toBe("2026-06-02T00:00:00Z");
  });

  test("accepts OOXML OnOff variants for w14:checked val (true / on / 1 / absent)", () => {
    // Word writes any of these forms; we previously only recognized "1"
    // and silently flipped state on save.
    const ns =
      'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"';
    expect(
      parseSdtPrXml(
        `<w:sdtPr ${ns}><w14:checkbox><w14:checked w14:val="true"/></w14:checkbox></w:sdtPr>`,
      ).checked,
    ).toBe(true);
    expect(
      parseSdtPrXml(
        `<w:sdtPr ${ns}><w14:checkbox><w14:checked w14:val="on"/></w14:checkbox></w:sdtPr>`,
      ).checked,
    ).toBe(true);
    expect(
      parseSdtPrXml(
        `<w:sdtPr ${ns}><w14:checkbox><w14:checked w14:val="1"/></w14:checkbox></w:sdtPr>`,
      ).checked,
    ).toBe(true);
    // A bare <w14:checked/> with no val attribute also means true per OnOff.
    expect(
      parseSdtPrXml(`<w:sdtPr ${ns}><w14:checkbox><w14:checked/></w14:checkbox></w:sdtPr>`).checked,
    ).toBe(true);
    // Negations.
    expect(
      parseSdtPrXml(
        `<w:sdtPr ${ns}><w14:checkbox><w14:checked w14:val="false"/></w14:checkbox></w:sdtPr>`,
      ).checked,
    ).toBe(false);
    expect(
      parseSdtPrXml(
        `<w:sdtPr ${ns}><w14:checkbox><w14:checked w14:val="0"/></w14:checkbox></w:sdtPr>`,
      ).checked,
    ).toBe(false);
  });

  test("listItem with only w:value falls back to value as displayText", () => {
    const props = parseSdtPrXml(
      '<w:sdtPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:dropDownList><w:listItem w:value="ca"/><w:listItem w:value="ny" w:displayText="New York"/></w:dropDownList></w:sdtPr>',
    );
    expect(props.sdtType).toBe("dropdown");
    expect(props.listItems).toEqual([
      { displayText: "ca", value: "ca" },
      { displayText: "New York", value: "ny" },
    ]);
  });

  test("listItem with only w:displayText falls back to displayText as value", () => {
    const props = parseSdtPrXml(
      '<w:sdtPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:dropDownList><w:listItem w:displayText="Yes"/></w:dropDownList></w:sdtPr>',
    );
    expect(props.listItems).toEqual([{ displayText: "Yes", value: "Yes" }]);
  });

  test("normalizes inherited alt-prefix w-namespace SDT children inside canonical wrapper", () => {
    // Canonical `<w:sdtPr>` wrapper with children inherited under an
    // alt prefix (`<x:tag>` declared via xmlns:x on the document root).
    // The pass-19 wrapper rewrite doesn't trigger because the wrapper
    // already uses `w:`, and the W14/W15 child fix doesn't apply
    // because `tag`/`alias` are w:-namespace names. Without W_LOCAL_NAMES
    // normalization the saved DOCX would carry undefined `x:` prefixes.
    const ns =
      'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:x="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
    const props = parseSdtPrXml(
      `<w:sdtPr ${ns}><x:tag x:val="client"/><x:alias x:val="Client Name"/></w:sdtPr>`,
    );
    expect(props.tag).toBe("client");
    expect(props.alias).toBe("Client Name");
    const written = rewrite(
      `<w:sdtPr ${ns}><x:tag x:val="client"/><x:alias x:val="Client Name"/></w:sdtPr>`,
    );
    expect(written).toContain('<w:tag w:val="client"/>');
    expect(written).toContain('<w:alias w:val="Client Name"/>');
    expect(written).not.toContain("x:tag");
    expect(written).not.toContain("x:alias");
  });

  test("leaves prefix-shaped substrings inside w:tag attribute values alone", () => {
    // A `w:tag` value carries opaque template-engine payloads in the wild
    // — most notably OpenDoPE conventions v2.3 (od:repeat=x1, od:xpath=…,
    // od:component=c1, etc.). The captured-raw normalization pass used to
    // rewrite ANY `\w+:` substring inside the `<w:tag …>` open tag,
    // including ones living inside the attribute *value*. That turned
    // `<w:tag w:val="od:repeat=x0"/>` into `<w:tag w:val="w:repeat=x0"/>`
    // on capture, corrupting every downstream consumer of `props.tag`
    // and breaking the parse → reconcile → re-parse round trip.
    const ns = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
    const props = parseSdtPrXml(`<w:sdtPr ${ns}><w:tag w:val="od:repeat=x0"/></w:sdtPr>`);
    expect(props.tag).toBe("od:repeat=x0");
    expect(rewrite(`<w:sdtPr ${ns}><w:tag w:val="od:repeat=x0"/></w:sdtPr>`)).toBe(
      '<w:sdtPr><w:tag w:val="od:repeat=x0"/></w:sdtPr>',
    );
  });

  test("does not rewrite prefix-shaped substrings inside w:tag attribute values containing whitespace", () => {
    // The previous regex-only attribute-name pass used `[^>]{0,200}\s` to
    // skip past the element's existing attribute heading. That match could
    // consume across the opening `"` of an attribute value and land its
    // anchor on whitespace *inside* the value, so a prefix-shaped token
    // sitting after that whitespace would be rewritten to canonical. A
    // `w:tag` whose value contains a single space before an OpenDoPE token
    // (a common shape in customer DOCX templates) triggered it.
    const ns = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
    // Whitespace inside value, prefix after the whitespace.
    expect(parseSdtPrXml(`<w:sdtPr ${ns}><w:tag w:val="foo od:repeat=x0"/></w:sdtPr>`).tag).toBe(
      "foo od:repeat=x0",
    );
    expect(rewrite(`<w:sdtPr ${ns}><w:tag w:val="foo od:repeat=x0"/></w:sdtPr>`)).toBe(
      '<w:sdtPr><w:tag w:val="foo od:repeat=x0"/></w:sdtPr>',
    );
    // Multiple prefix-shaped tokens separated by whitespace inside the
    // value. Each one was a potential rewrite anchor under the old pass.
    expect(parseSdtPrXml(`<w:sdtPr ${ns}><w:tag w:val="a:b c:d"/></w:sdtPr>`).tag).toBe("a:b c:d");
    expect(rewrite(`<w:sdtPr ${ns}><w:tag w:val="a:b c:d"/></w:sdtPr>`)).toBe(
      '<w:sdtPr><w:tag w:val="a:b c:d"/></w:sdtPr>',
    );
  });

  test("leaves nested rPr color elements alone (no overzealous w15 rewrite)", () => {
    // A placeholder rPr inside sdtPr carries `<w:color w:val="…"/>` for
    // run text color. That local name is `color` but it lives in the
    // w: namespace, not w15. The earlier local-name rewrite would have
    // re-emitted it under w15, corrupting run formatting on every
    // parse → save round trip.
    const ns = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
    const source = `<w:sdtPr ${ns}><w:rPr><w:color w:val="FF0000"/></w:rPr><w:tag w:val="x"/></w:sdtPr>`;
    expect(parseSdtPrXml(source).tag).toBe("x");
    // Color stays under w:, not silently rewritten to w15:, and `w:rPr`
    // leads the sequence the way `CT_SdtPr` declares it.
    expect(rewrite(source)).toBe(
      '<w:sdtPr><w:rPr><w:color w:val="FF0000"/></w:rPr><w:tag w:val="x"/></w:sdtPr>',
    );
  });

  test("an inherited alt-prefix extension child replays with its own binding", () => {
    // Source binds the w14 / w15 URIs under non-canonical prefixes at the
    // document root and the children inherit `x:` / `y:`. Each capture
    // materialises the binding it uses, so the fragment is self-contained
    // wherever it is written back: the saved part never carries an
    // unresolved prefix, whatever the root happens to declare.
    const ns =
      'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:x="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:y="http://schemas.microsoft.com/office/word/2012/wordml"';
    const source = `<w:sdtPr ${ns}><x:checkbox><x:checked x:val="1"/></x:checkbox><y:repeatingSection/></w:sdtPr>`;
    const props = parseSdtPrXml(source);
    expect(props.sdtType).toBe("checkbox");
    expect(props.checked).toBe(true);
    expect(rewrite(source)).toBe(
      '<w:sdtPr><x:checkbox xmlns:x="http://schemas.microsoft.com/office/word/2010/wordml">' +
        '<x:checked x:val="1"/></x:checkbox>' +
        '<y:repeatingSection xmlns:y="http://schemas.microsoft.com/office/word/2012/wordml"/></w:sdtPr>',
    );
  });

  test("an alt-prefix property set is written back canonically", () => {
    // A producer that binds the WordprocessingML namespace under `ns0`
    // emits `<ns0:sdtPr>…</ns0:sdtPr>`. Every child of it is modelled, so
    // the writer rebuilds them under the canonical prefix — and in the
    // order `CT_SdtPr` declares rather than the order the source wrote.
    const ns =
      'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:ns0="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
    const source = `<ns0:sdtPr ${ns}><ns0:tag ns0:val="client"/><ns0:alias ns0:val="Client"/></ns0:sdtPr>`;
    const props = parseSdtPrXml(source);
    expect(props.tag).toBe("client");
    expect(props.alias).toBe("Client");
    expect(rewrite(source)).toBe(
      '<w:sdtPr><w:alias w:val="Client"/><w:tag w:val="client"/></w:sdtPr>',
    );
  });

  test("reads placeholder docPart reference from the docPart's own w:val attribute", () => {
    // The parser used to look for a nested `<w:val>` child inside
    // `<w:docPart>`. The actual OOXML shape is `<w:docPart w:val="…"/>`,
    // so the placeholder reference was silently dropped.
    const ns = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
    const props = parseSdtPrXml(
      `<w:sdtPr ${ns}><w:placeholder><w:docPart w:val="DefaultPlaceholder"/></w:placeholder></w:sdtPr>`,
    );
    expect(props.placeholder).toBe("DefaultPlaceholder");
  });

  test("reads top-level SDT property attrs under an alternative namespace prefix", () => {
    // Same prefix-tolerance theme as listItems and checkbox val: the
    // top-level w:tag / w:alias / w:id / w:lock readers used to be hard-
    // coded to the `w:` prefix, so docs that bound the namespace under
    // ns0 came back with undefined tag / alias / lock and could not be
    // addressed via getContentControls({ tag }).
    const ns =
      'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:ns0="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
    const props = parseSdtPrXml(
      `<w:sdtPr ${ns}><ns0:id ns0:val="42"/><ns0:tag ns0:val="client"/><ns0:alias ns0:val="Client Name"/><ns0:lock ns0:val="contentLocked"/></w:sdtPr>`,
    );
    expect(props.id).toBe(42);
    expect(props.tag).toBe("client");
    expect(props.alias).toBe("Client Name");
    expect(props.lock).toBe("contentLocked");
  });

  test("reads dropdown listItem attrs under an alternative namespace prefix", () => {
    // Same prefix-tolerance bug as checkbox val — `parseListItems` matched
    // `<ns0:listItem>` via local-name fallback but then read `w:displayText`
    // / `w:value`, which came back null and the item was silently dropped.
    // A dropdown bound under any non-`w` prefix would open with no options.
    const ns =
      'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:ns0="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
    const props = parseSdtPrXml(
      `<w:sdtPr ${ns}><w:dropDownList><ns0:listItem ns0:displayText="A" ns0:value="a"/><ns0:listItem ns0:displayText="B" ns0:value="b"/></w:dropDownList></w:sdtPr>`,
    );
    expect(props.sdtType).toBe("dropdown");
    expect(props.listItems).toEqual([
      { displayText: "A", value: "a" },
      { displayText: "B", value: "b" },
    ]);
  });

  test("reads checkbox val under an alternative namespace prefix", () => {
    // OOXML binds prefixes to URIs at the document root; a valid doc can
    // bind the w14 namespace under any prefix (here `ns0`). The marker
    // element matched fine via local-name fallback, but the val attribute
    // was missed because we only looked for `w14:val` / bare `val`. An
    // unchecked box was then parsed as checked.
    const ns =
      'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:ns0="http://schemas.microsoft.com/office/word/2010/wordml"';
    expect(
      parseSdtPrXml(
        `<w:sdtPr ${ns}><ns0:checkbox><ns0:checked ns0:val="0"/></ns0:checkbox></w:sdtPr>`,
      ).checked,
    ).toBe(false);
    expect(
      parseSdtPrXml(
        `<w:sdtPr ${ns}><ns0:checkbox><ns0:checked ns0:val="1"/></ns0:checkbox></w:sdtPr>`,
      ).checked,
    ).toBe(true);
    expect(
      parseSdtPrXml(
        `<w:sdtPr ${ns}><ns0:checkbox><ns0:checked ns0:val="false"/></ns0:checkbox></w:sdtPr>`,
      ).checked,
    ).toBe(false);
  });

  test("respects w:showingPlcHdr OnOff val (false / off / 0)", () => {
    // The presence-implies-true semantics still applies, but an explicit
    // negation must be honored. Previously the parser flipped
    // `val="false"` back to `true` because it ignored the attribute.
    const ns = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
    expect(parseSdtPrXml(`<w:sdtPr ${ns}><w:showingPlcHdr/></w:sdtPr>`).showingPlaceholder).toBe(
      true,
    );
    expect(
      parseSdtPrXml(`<w:sdtPr ${ns}><w:showingPlcHdr w:val="true"/></w:sdtPr>`).showingPlaceholder,
    ).toBe(true);
    expect(
      parseSdtPrXml(`<w:sdtPr ${ns}><w:showingPlcHdr w:val="false"/></w:sdtPr>`).showingPlaceholder,
    ).toBe(false);
    expect(
      parseSdtPrXml(`<w:sdtPr ${ns}><w:showingPlcHdr w:val="0"/></w:sdtPr>`).showingPlaceholder,
    ).toBe(false);
    expect(
      parseSdtPrXml(`<w:sdtPr ${ns}><w:showingPlcHdr w:val="off"/></w:sdtPr>`).showingPlaceholder,
    ).toBe(false);
  });

  test("falls back to richText when the marker is unknown", () => {
    const props = parseSdtPrXml(
      '<w:sdtPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"><w15:appearance w15:val="boundingBox"/><w:tag w:val="t"/></w:sdtPr>',
    );
    expect(props.sdtType).toBe("richText");
    expect(props.tag).toBe("t");
  });

  test("writes modelled children in CT_SdtPr declaration order", () => {
    expect(
      serializeSdtProperties({
        sdtType: "date",
        alias: "Date",
        tag: "date-field",
        id: 42,
        lock: "contentLocked",
        placeholder: "DatePlaceholder",
        showingPlaceholder: true,
        dateFormat: "yyyy-MM-dd",
      }),
    ).toBe(
      '<w:sdtPr><w:alias w:val="Date"/><w:tag w:val="date-field"/><w:id w:val="42"/>' +
        '<w:lock w:val="contentLocked"/><w:placeholder><w:docPart w:val="DatePlaceholder"/>' +
        '</w:placeholder><w:showingPlcHdr/><w:date><w:dateFormat w:val="yyyy-MM-dd"/>' +
        "</w:date></w:sdtPr>",
    );
  });

  test("rejects malformed preserved kind markup before patching modelled state", () => {
    const mismatchedDate =
      '<w:date xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:x="urn:other"><w:dateFormat w:val="old"/></x:date>';

    expect(() =>
      serializeSdtProperties({
        sdtType: "date",
        dateFormat: "yyyy-MM-dd",
        preserved: { children: [{ index: 19, xml: mismatchedDate }] },
      }),
    ).toThrow("Preserved container markup must be one bounded, well-formed XML element");
  });
});
