import { describe, expect, test } from "bun:test";

import {
  OOXML_NAMESPACES,
  UnboundNamespacePrefixError,
  readRootNamespaceBindings,
  serializePartElement,
} from "./partNamespaces";

const declaredPrefixes = (xml: string): string[] =>
  [...xml.matchAll(/xmlns:(?<prefix>[\w.-]+)=/gu)].map((match) => match.groups?.prefix ?? "");

const ignorablePrefixes = (xml: string): string[] => {
  const attribute = /mc:Ignorable="(?<value>[^"]*)"/u.exec(xml);
  return attribute ? (attribute.groups?.value ?? "").split(" ").filter(Boolean) : [];
};

const serialize = (body: string, sourceBindings?: ReadonlyMap<string, string>): string =>
  serializePartElement({
    partPath: "word/document.xml",
    rootName: "w:document",
    baselinePrefixes: ["mc", "w", "w14"],
    sourceBindings,
    body,
  });

describe("serializePartElement", () => {
  test("declares the baseline even when the body uses none of it", () => {
    expect(declaredPrefixes(serialize("<w:body/>"))).toEqual(["mc", "w", "w14"]);
  });

  test("declares a prefix the body uses beyond the baseline", () => {
    const xml = serialize("<w:body><wp:txbx><wne:txbxContent/></wp:txbx></w:body>");
    expect(declaredPrefixes(xml)).toContain("wne");
    expect(xml).toContain(`xmlns:wne="${OOXML_NAMESPACES.wne.uri}"`);
  });

  test("declares a prefix carried only by an attribute name", () => {
    expect(
      declaredPrefixes(serialize('<w:body><w:p w16du:dateUtc="2024-01-01T00:00:00Z"/></w:body>')),
    ).toContain("w16du");
  });

  test("keeps the URI the source part bound for a prefix the table does not know", () => {
    const xml = serialize(
      "<w:body><acme:marker/></w:body>",
      new Map([["acme", "urn:acme:markers"]]),
    );
    expect(xml).toContain('xmlns:acme="urn:acme:markers"');
  });

  test("prefers the table over a source binding that disagrees", () => {
    const xml = serialize("<w:body/>", new Map([["w", "urn:wrong"]]));
    expect(xml).toContain(`xmlns:w="${OOXML_NAMESPACES.w.uri}"`);
    expect(xml).not.toContain("urn:wrong");
  });

  test("leaves a prefix the body binds for its own subtree to that binding", () => {
    const body = '<w:body><a14:x xmlns:a14="urn:inline"/></w:body>';
    expect(declaredPrefixes(serialize(body))).toEqual(["mc", "w", "w14", "a14"]);
  });

  test("fails on a prefix nothing binds rather than writing it unbound", () => {
    expect(() => serialize("<w:body><acme:marker/></w:body>")).toThrow(UnboundNamespacePrefixError);
  });

  test("reports every unbound prefix in the part", () => {
    try {
      serialize("<w:body><acme:a/><zeta:b/></w:body>");
      expect.unreachable();
    } catch (error) {
      expect(error).toBeInstanceOf(UnboundNamespacePrefixError);
      expect((error as UnboundNamespacePrefixError).prefixes).toEqual(["acme", "zeta"]);
      expect((error as UnboundNamespacePrefixError).partPath).toBe("word/document.xml");
    }
  });

  test("a colon inside an attribute value is not a namespace prefix", () => {
    const xml = serialize('<w:body><w:p w:rsidR="see http://x/y note:this"/></w:body>');
    expect(declaredPrefixes(xml)).toEqual(["mc", "w", "w14"]);
  });

  test("comments, CDATA and processing instructions carry no prefixes", () => {
    const body = "<w:body><!-- acme:not-a-tag --><![CDATA[<zeta:nope/>]]><?pi acme:x?></w:body>";
    expect(declaredPrefixes(serialize(body))).toEqual(["mc", "w", "w14"]);
  });

  test("mc:Ignorable lists exactly the declared prefixes marked ignorable", () => {
    const xml = serialize('<w:body><w:p w16du:dateUtc="2024-01-01T00:00:00Z"/></w:body>');
    const declared = new Set(declaredPrefixes(xml));
    const listed = ignorablePrefixes(xml);
    expect(listed).toEqual(["w14", "w16du"]);
    for (const prefix of listed) {
      expect(declared.has(prefix)).toBe(true);
    }
    for (const prefix of declared) {
      const entry = Object.entries(OOXML_NAMESPACES).find(([name]) => name === prefix)?.[1];
      expect(entry?.ignorable === true).toBe(listed.includes(prefix));
    }
  });

  test("omits mc:Ignorable when no declared prefix is ignorable", () => {
    const xml = serializePartElement({
      partPath: "word/styles.xml",
      rootName: "w:styles",
      baselinePrefixes: ["w"],
      sourceBindings: undefined,
      body: "<w:style/>",
    });
    expect(xml).not.toContain("mc:Ignorable");
    expect(xml).not.toContain("xmlns:mc");
  });

  test("keeps root attributes that are not namespace declarations", () => {
    const xml = serializePartElement({
      partPath: "word/theme/theme1.xml",
      rootName: "a:theme",
      rootAttributes: 'name="Folio"',
      baselinePrefixes: ["a"],
      sourceBindings: undefined,
      body: "<a:themeElements/>",
    });
    expect(xml.startsWith('<a:theme name="Folio" xmlns:a=')).toBe(true);
  });
});

describe("readRootNamespaceBindings", () => {
  test("reads the root element's declarations past the prolog", () => {
    const bindings = readRootNamespaceBindings(
      '<?xml version="1.0"?>\n<!-- a note -->\n' +
        '<w:document xmlns:w="urn:w" xmlns:acme="urn:acme"><w:body><x:y xmlns:x="urn:inner"/></w:body></w:document>',
    );
    expect([...bindings.entries()].sort()).toEqual([
      ["acme", "urn:acme"],
      ["w", "urn:w"],
    ]);
  });

  test("is empty for markup with no root declarations", () => {
    expect(readRootNamespaceBindings("<w:hdr><w:p/></w:hdr>").size).toBe(0);
  });
});
