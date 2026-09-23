import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import JSZip from "jszip";
import { readFileSync } from "node:fs";
import { resolve } from "node:path";

import { propertyTestTimeout } from "../../../../test/property-testing";

import { parseStreamingXml, rewriteStreamingXmlDecimalAttributes } from "./streamingXmlParser";
import {
  getNamespaceUri,
  OOXML_NAMESPACE_SCOPE,
  parseXml,
  parseXmlWithFastXmlParser,
  type XmlElement,
  type XmlNamespaceScope,
} from "./xmlParser";

setDefaultTimeout(propertyTestTimeout(30_000));

const REPO_ROOT = resolve(import.meta.dir, "../../../..");
const DOCUMENT_FIXTURE_GLOBS = [
  "tests/visual/fixtures/*.docx",
  "packages/core/src/docx/__fixtures__/*.docx",
  "packages/core/src/docx/__tests__/__fixtures__/**/*.docx",
] as const;

/** Every binding in scope, flattened, so two chains that resolve alike compare equal. */
const scopeKey = (() => {
  const keys = new WeakMap<XmlNamespaceScope, string>();
  const key = (scope: XmlNamespaceScope | undefined): string => {
    if (scope === undefined) {
      return "[]";
    }
    const cached = keys.get(scope);
    if (cached !== undefined) {
      return cached;
    }
    const resolved = new Map<string, string>(JSON.parse(key(scope.parent)));
    for (const [prefix, uri] of scope.bindings) {
      resolved.set(prefix, uri);
    }
    const computed = JSON.stringify(
      [...resolved].toSorted(([left], [right]) => left.localeCompare(right)),
    );
    keys.set(scope, computed);
    return computed;
  };
  return key;
})();

/**
 * The tree with each element's resolved namespace and in-scope bindings,
 * which `toEqual` cannot see: both are non-enumerable.
 */
const describeTree = (root: XmlElement): string =>
  JSON.stringify(root, (key, value: XmlElement) =>
    key !== "" && value?.type === "element"
      ? { ...value, namespace: getNamespaceUri(value), scope: scopeKey(value.namespaceScope) }
      : value,
  );

/**
 * fast-xml-parser keeps a byte-order mark ahead of the declaration as a text
 * node beside the root element. No reader consults text outside the root, and
 * the streaming reader does not produce it.
 */
const withoutRootWhitespace = (root: XmlElement): XmlElement => ({
  ...root,
  elements: (root.elements ?? []).filter(
    (node) => node.type !== "text" || String(node.text).trim() !== "",
  ),
});

const expectStreamingMatchesFallback = (
  xml: string,
  label: string,
  scope?: XmlNamespaceScope,
): void => {
  const streaming = parseStreamingXml(xml, scope);
  expect(streaming.status, label).toBe("parsed");
  if (streaming.status !== "parsed") {
    return;
  }
  const fallback = withoutRootWhitespace(parseXmlWithFastXmlParser(xml, scope));
  expect(describeTree(streaming.value), label).toBe(describeTree(fallback));
};

const XML_CASES = [
  `<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="urn:w">
  <!-- ignored -->
  <w:body>
    <w:p data-one="&quot;&amp;&#65;&#x42;" data-two='apostrophe'>
      <w:r><w:t xml:space="preserve"> text &lt; value </w:t></w:r>
    </w:p>
  </w:body>
</w:document>`,
  "<document>\r\n  <body><p><![CDATA[raw <xml>\r\n&\rtext]]></p></body>\r</document>",
  "<document><text><![CDATA[first]]><!-- ignored -->second</text></document>",
  "<x:document><x:body><x:p/><x:tbl><x:tr><x:tc/></x:tr></x:tbl></x:body></x:document>",
] as const;

describe("parseStreamingXml", () => {
  test.each(XML_CASES)("matches the compatibility parser", (xml) => {
    expect(parseStreamingXml(xml)).toEqual({
      status: "parsed",
      value: parseXmlWithFastXmlParser(xml),
    });
  });

  test("falls back for declarations and malformed nesting outside its contract", () => {
    const deeplyNested = `<document>${"<x>".repeat(101)}value${"</x>".repeat(101)}</document>`;

    expect(parseStreamingXml("<!DOCTYPE document><document/>").status).toBe("unsupported");
    expect(parseStreamingXml("<document><body></document>").status).toBe("unsupported");
    expect(parseStreamingXml("<document>&custom;</document>").status).toBe("unsupported");
    expect(parseStreamingXml('<document __proto__="unsafe"/>').status).toBe("unsupported");
    expect(parseStreamingXml(deeplyNested).status).toBe("unsupported");
  });

  test("matches inherited and rebound namespace metadata", () => {
    const xml = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:body>
        <w:p w:id="transitional"><w:r/></w:p>
        <x:p xmlns:x="http://purl.oclc.org/ooxml/wordprocessingml/main" x:id="strict"><x:r/></x:p>
        <w:p xmlns:w="https://example.com/foreign" w:id="foreign"><w:r/></w:p>
        <w:p w:id="restored"><w:r/></w:p>
      </w:body>
    </w:document>`;
    expectStreamingMatchesFallback(xml, "rebound namespaces");
  });

  test("resolves a fragment against the scope it was captured under", () => {
    const fragment = `<w:pPr><w:jc w:val="center"/><m:oMathPara/></w:pPr>`;
    expectStreamingMatchesFallback(fragment, "fragment", OOXML_NAMESPACE_SCOPE);
    const parsed = parseXml(fragment, OOXML_NAMESPACE_SCOPE).elements?.at(0);
    expect(parsed === undefined ? undefined : getNamespaceUri(parsed)).toBe(
      "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    );
  });

  test("parseXml falls back for markup outside the streaming contract", () => {
    const xml = "<document>&nbsp;</document>";
    expect(parseStreamingXml(xml).status).toBe("unsupported");
    expect(parseXml(xml)).toEqual(parseXmlWithFastXmlParser(xml));
  });

  test("matches entity and line-ending behavior for generated values", () => {
    const encodedToken = fc.constantFrom(
      "text",
      " ",
      "\t",
      "\r",
      "\r\n",
      "žluťoučký",
      "日本語",
      "&amp;",
      "&apos;",
      "&gt;",
      "&lt;",
      "&quot;",
      "&#65;",
      "&#x1F642;",
    );

    fc.assert(
      fc.property(
        fc.array(encodedToken, { maxLength: 40 }),
        fc.array(encodedToken, { maxLength: 20 }),
        (textTokens, attributeTokens) => {
          const xml = `<document value="${attributeTokens.join("")}"><text>${textTokens.join(
            "",
          )}</text><empty/></document>`;
          expect(parseStreamingXml(xml)).toEqual({
            status: "parsed",
            value: parseXmlWithFastXmlParser(xml),
          });
        },
      ),
      { numRuns: 250 },
    );
  });

  test("matches the compatibility parser for generated trees", () => {
    const text = fc
      .array(fc.constantFrom("a", " ", "\n", "\r\n", "\t", "ž", "&amp;", "&lt;", "&#x1F642;"), {
        maxLength: 6,
      })
      .map((tokens) => tokens.join(""));
    const name = fc.constantFrom("w:p", "w:r", "m:r", "x:p", "p", "w:t");
    const attribute = fc.oneof(
      fc.tuple(fc.constantFrom("w:val", "x:id", "id", "xml:space"), text),
      fc.tuple(
        fc.constantFrom("xmlns:w", "xmlns:x", "xmlns:m", "xmlns"),
        fc.constantFrom(
          "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
          "http://purl.oclc.org/ooxml/wordprocessingml/main",
          "http://schemas.openxmlformats.org/officeDocument/2006/math",
          "urn:foreign",
        ),
      ),
    );
    const { node } = fc.letrec<{ node: string }>((tie) => ({
      node: fc.oneof(
        { depthSize: "small", withCrossShrink: true },
        text,
        fc.constantFrom("<!-- note -->", "<![CDATA[raw <x> & ]]>", "<?pi data?>"),
        fc
          .tuple(
            name,
            fc.uniqueArray(attribute, { selector: ([key]) => key, maxLength: 3 }),
            fc.array(tie("node"), { maxLength: 4 }),
          )
          .map(([tag, attributes, children]) => {
            const attrs = attributes.map(([key, value]) => ` ${key}="${value}"`).join("");
            return children.length === 0
              ? `<${tag}${attrs}/>`
              : `<${tag}${attrs}>${children.join("")}</${tag}>`;
          }),
      ),
    }));
    const document = fc
      .tuple(fc.array(node, { maxLength: 5 }), fc.constantFrom(undefined, OOXML_NAMESPACE_SCOPE))
      .map(([children, scope]) => ({
        xml: `<?xml version="1.0"?>\n<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${children.join("")}</w:document>`,
        scope,
      }));

    fc.assert(
      fc.property(document, ({ xml, scope }) => {
        expectStreamingMatchesFallback(xml, xml, scope);
      }),
      { numRuns: 300 },
    );
  });

  test("matches the compatibility parser for every XML part in the repository fixtures", async () => {
    const paths: string[] = [];
    for (const pattern of DOCUMENT_FIXTURE_GLOBS) {
      paths.push(...new Bun.Glob(pattern).scanSync({ cwd: REPO_ROOT }));
    }
    expect(paths.length).toBeGreaterThan(30);

    for (const path of paths) {
      const zip = await JSZip.loadAsync(readFileSync(resolve(REPO_ROOT, path)));
      expect(zip.file("word/document.xml"), path).not.toBeNull();
      for (const file of zip.file(/\.(?:xml|rels)$/iu)) {
        expectStreamingMatchesFallback(await file.async("string"), `${path}:${file.name}`);
      }
    }
  });
});

test.each(["' injected='yes", '"/>', "&quot;", "1.5", "-1"])(
  "refuses a non-decimal attribute replacement: %s",
  (replacement) => {
    expect(
      rewriteStreamingXmlDecimalAttributes('<r id="1"/>', () => new Map([["id", replacement]])),
    ).toEqual({ status: "unsupported" });
  },
);
