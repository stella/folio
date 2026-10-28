import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import {
  getXmlSourceRange,
  parseStreamingXml,
  parseStreamingXmlWithSourceRanges,
  rewriteStreamingXmlDecimalAttributes,
  scanStreamingXmlNumericIdAttributes,
  XmlReadError,
} from "@stll/docx-core/xml";

import {
  propertyConfig,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import {
  elementToXml,
  getNamespaceUri,
  OOXML_NAMESPACE_SCOPE,
  parseXml,
  type XmlElement,
  type XmlNamespaceScope,
} from "./xmlParser";

setDefaultTimeout(propertyTestTimeout(30_000));

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
/** XML 1.0 §2.11 and §4.6, restated for the tokens the generators draw. */
const expectedValue = (raw: string): string =>
  raw
    .replaceAll("\r\n", "\n")
    .replaceAll("\r", "\n")
    .replace(
      /&(amp|apos|gt|lt|quot|#[0-9]+|#x[0-9A-Fa-f]+);/gu,
      (_, entity: string) => {
        switch (entity) {
          case "amp":
            return "&";
          case "apos":
            return "'";
          case "gt":
            return ">";
          case "lt":
            return "<";
          case "quot":
            return '"';
          default:
            return String.fromCodePoint(
              entity.startsWith("#x")
                ? Number.parseInt(entity.slice(2), 16)
                : Number.parseInt(entity.slice(1), 10),
            );
        }
      },
    );

const describeTree = (root: XmlElement): string =>
  JSON.stringify(root, (key, value: XmlElement) =>
    key !== "" && value?.type === "element"
      ? {
          ...value,
          namespace: getNamespaceUri(value),
          scope: scopeKey(value.namespaceScope),
        }
      : value,
  );

/**
 * Each element's namespace URI, resolved independently of the reader: the
 * nearest declaration of its prefix on itself or an ancestor, else the scope
 * the fragment was parsed under.
 */
const expectNamespacesResolved = (
  element: XmlElement,
  inherited: ReadonlyMap<string, string>,
): void => {
  if (element.type !== "element") return;
  const bindings = new Map(inherited);
  for (const [name, value] of Object.entries(element.attributes ?? {})) {
    if (name === "xmlns") bindings.set("", String(value));
    else if (name.startsWith("xmlns:"))
      bindings.set(name.slice("xmlns:".length), String(value));
  }
  const name = element.name ?? "";
  const colon = name.indexOf(":");
  expect(getNamespaceUri(element), name).toBe(
    bindings.get(colon === -1 ? "" : name.slice(0, colon)),
  );
  for (const child of element.elements ?? []) {
    expectNamespacesResolved(child, bindings);
  }
};

describe("parseStreamingXml", () => {
  test("source ranges are opt-in and exact across namespace profiles and nested containers", () => {
    fc.assert(
      fc.property(
        fc.constantFrom("w", "alias", ""),
        fc.constantFrom(
          "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
          "http://purl.oclc.org/ooxml/wordprocessingml/main",
        ),
        fc.constantFrom(" ", "\n", "\r\n"),
        fc.constantFrom('"', "'"),
        (prefix, uri, whitespace, quote) => {
          const name = (local: string) =>
            prefix ? `${prefix}:${local}` : local;
          const namespace = prefix ? `xmlns:${prefix}` : "xmlns";
          const text = `<${name("t")}>é😀<!--nested comment--><![CDATA[<raw>&]]></${name("t")}>`;
          const run = `<${name("r")}>${text}</${name("r")}>`;
          const paragraph = `<${name("p")} data=${quote}a>b${quote}>${run}</${name("p")}>`;
          const empty = `<${name("p")}${whitespace}/>`;
          const cell = `<${name("tc")}>${paragraph}${whitespace}${empty}</${name("tc")}>`;
          const row = `<${name("tr")}>${cell}</${name("tr")}>`;
          const table = `<${name("tbl")}>${row}</${name("tbl")}>`;
          const rebound = `<${name("other")} ${namespace}=${quote}urn:other${quote}/>`;
          const root = `<${name("hdr")} ${namespace}=${quote}${uri}${quote}>${whitespace}<!--root comment-->${table}${rebound}</${name("hdr")}>`;
          const xml = `<?xml version="1.0"?>${whitespace}${root}${whitespace}`;
          const ordinary = parseStreamingXml(xml);
          const tracked = parseStreamingXmlWithSourceRanges(xml);
          expect(ordinary.status).toBe("parsed");
          expect(tracked.status).toBe("parsed");
          if (ordinary.status !== "parsed" || tracked.status !== "parsed")
            return;
          expect(describeTree(tracked.value)).toBe(
            describeTree(ordinary.value),
          );
          const visitOrdinary = (element: XmlElement): void => {
            expect(getXmlSourceRange(element)).toBeUndefined();
            for (const child of element.elements ?? []) visitOrdinary(child);
          };
          visitOrdinary(ordinary.value);
          const expected = [
            root,
            table,
            row,
            cell,
            paragraph,
            run,
            text,
            empty,
            rebound,
          ];
          const actual: string[] = [];
          const visitTracked = (element: XmlElement): void => {
            if (element.type === "element") {
              const range = getXmlSourceRange(element);
              expect(range).toBeDefined();
              if (range) actual.push(xml.slice(range.start, range.end));
              expect(getNamespaceUri(element)).toBe(
                element.name === name("other") ? "urn:other" : uri,
              );
            } else {
              expect(getXmlSourceRange(element)).toBeUndefined();
            }
            for (const child of element.elements ?? []) visitTracked(child);
          };
          visitTracked(tracked.value);
          expect(actual).toEqual(expected);
        },
      ),
      propertyConfig(),
    );
  });
  test("identity scans preserve namespace and source spans without retaining content", () => {
    fc.assert(
      fc.property(
        fc.integer({ min: 0, max: 2_147_483_647 }),
        fc.array(fc.constantFrom("text", "&amp;", "&#65;", "日本語"), {
          maxLength: 100,
        }),
        (id, text) => {
          const namespace =
            "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
          const xml = `<x:document xmlns:x="${namespace}"><x:p data-content="discard"><x:commentRangeStart x:id="${id}"/>${text.join("")}<x:p xmlns:x="urn:foreign"><x:commentReference x:id='${id}'/></x:p><x:commentRangeEnd x:id="${id}"/></x:p></x:document>`;
          const events: {
            name: string | undefined;
            namespace: string | undefined;
            id: string;
          }[] = [];
          const result = scanStreamingXmlNumericIdAttributes(
            xml,
            (element, spans) => {
              expect(element.elements).toBeUndefined();
              expect(element.attributes?.["data-content"]).toBeUndefined();
              const span = spans.get("x:id");
              if (span) {
                events.push({
                  name: element.name,
                  namespace: getNamespaceUri(element),
                  id: xml.slice(span.start, span.end),
                });
              }
              return null;
            },
          );
          expect(result).toEqual({ status: "scanned" });
          expect(events).toEqual([
            { name: "x:commentRangeStart", namespace, id: String(id) },
            {
              name: "x:commentReference",
              namespace: "urn:foreign",
              id: String(id),
            },
            { name: "x:commentRangeEnd", namespace, id: String(id) },
          ]);
        },
      ),
      propertyConfig({ numRuns: 100 }),
    );
  });

  test.each([
    "<!DOCTYPE root><root/>",
    "<root><child></root>",
    "<root>&custom;</root>",
    '<root unrelated="&custom;"/>',
    '<root __proto__="unsafe"/>',
    `<root>${"<x>".repeat(101)}value${"</x>".repeat(101)}</root>`,
  ])("identity scans reject unsupported XML: %s", (xml) => {
    expect(scanStreamingXmlNumericIdAttributes(xml, () => null)).toEqual({
      status: "unsupported",
    });
  });

  test("refuses declarations and malformed nesting outside its contract", () => {
    const deeplyNested = `<document>${"<x>".repeat(101)}value${"</x>".repeat(101)}</document>`;

    expect(parseStreamingXml("<!DOCTYPE document><document/>").status).toBe(
      "unsupported",
    );
    expect(parseStreamingXml("<document><body></document>").status).toBe(
      "unsupported",
    );
    expect(parseStreamingXml("<document>&custom;</document>").status).toBe(
      "unsupported",
    );
    expect(parseStreamingXml('<document __proto__="unsafe"/>').status).toBe(
      "unsupported",
    );
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
    const parsed = parseStreamingXml(xml);
    expect(parsed.status).toBe("parsed");
    if (parsed.status !== "parsed") return;
    const paragraphs = (parsed.value.elements?.at(0)?.elements ?? [])
      .flatMap((node) => node.elements ?? [])
      .filter((node) => node.type === "element");
    expect(
      paragraphs.map((paragraph) => [
        paragraph.attributes?.["w:id"] ?? paragraph.attributes?.["x:id"],
        getNamespaceUri(paragraph),
        getNamespaceUri(paragraph.elements?.at(0) ?? {}),
      ]),
    ).toEqual([
      [
        "transitional",
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
      ],
      [
        "strict",
        "http://purl.oclc.org/ooxml/wordprocessingml/main",
        "http://purl.oclc.org/ooxml/wordprocessingml/main",
      ],
      ["foreign", "https://example.com/foreign", "https://example.com/foreign"],
      [
        "restored",
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
      ],
    ]);
  });

  test("resolves a fragment against the scope it was captured under", () => {
    const fragment = `<w:pPr><w:jc w:val="center"/><m:oMathPara/></w:pPr>`;
    const streamed = parseStreamingXml(fragment, OOXML_NAMESPACE_SCOPE);
    expect(streamed.status).toBe("parsed");
    if (streamed.status === "parsed") {
      expectNamespacesResolved(
        streamed.value.elements?.at(0) ?? {},
        OOXML_NAMESPACE_SCOPE.bindings,
      );
    }
    const parsed = parseXml(fragment, OOXML_NAMESPACE_SCOPE).elements?.at(0);
    expect(parsed === undefined ? undefined : getNamespaceUri(parsed)).toBe(
      "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    );
  });

  test("parseXml reads the references earlier releases read, and refuses broken markup", () => {
    const xml = "<document>&nbsp;&custom;&#0;</document>";
    expect(parseStreamingXml(xml).status).toBe("unsupported");
    expect(parseXml(xml).elements?.at(0)?.elements).toEqual([
      { type: "text", text: "\u00a0&custom;" },
    ]);
    expect(() => parseXml("<document><body></document>")).toThrow(XmlReadError);
    expect(() => parseXml("<!DOCTYPE document><document/>")).toThrow(
      XmlReadError,
    );
  });

  test("decodes references and line endings in generated values", () => {
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
          const parsed = parseStreamingXml(xml);
          expect(parsed.status).toBe("parsed");
          if (parsed.status !== "parsed") return;
          const document = parsed.value.elements?.at(0);
          expect(document?.attributes?.["value"]).toBe(
            expectedValue(attributeTokens.join("")),
          );
          const text = expectedValue(textTokens.join(""));
          expect(document?.elements?.at(0)?.elements).toEqual(
            text === "" ? undefined : [{ type: "text", text }],
          );
        },
      ),
      propertyConfig({ numRuns: 250 }),
    );
  });

  test("generated trees resolve namespaces and survive a write and a re-read", () => {
    const text = fc
      .array(
        fc.constantFrom(
          "a",
          " ",
          "\n",
          "\r\n",
          "\t",
          "ž",
          "&amp;",
          "&lt;",
          "&#x1F642;",
        ),
        {
          maxLength: 6,
        },
      )
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
        fc.constantFrom(
          "<!-- note -->",
          "<![CDATA[raw <x> & ]]>",
          "<?pi data?>",
        ),
        fc
          .tuple(
            name,
            fc.uniqueArray(attribute, {
              selector: ([key]) => key,
              maxLength: 3,
            }),
            fc.array(tie("node"), { maxLength: 4 }),
          )
          .map(([tag, attributes, children]) => {
            const attrs = attributes
              .map(([key, value]) => ` ${key}="${value}"`)
              .join("");
            return children.length === 0
              ? `<${tag}${attrs}/>`
              : `<${tag}${attrs}>${children.join("")}</${tag}>`;
          }),
      ),
    }));
    const document = fc
      .tuple(
        fc.array(node, { maxLength: 5 }),
        fc.constantFrom(undefined, OOXML_NAMESPACE_SCOPE),
      )
      .map(([children, scope]) => ({
        xml: `<?xml version="1.0"?>\n<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${children.join("")}</w:document>`,
        scope,
      }));

    fc.assert(
      fc.property(document, ({ xml, scope }) => {
        const parsed = parseStreamingXml(xml, scope);
        expect(parsed.status, xml).toBe("parsed");
        if (parsed.status !== "parsed") return;
        const root = parsed.value.elements?.find(
          ({ type }) => type === "element",
        );
        expect(root).toBeDefined();
        if (root === undefined) return;
        expectNamespacesResolved(root, scope?.bindings ?? new Map());
        const reread = parseStreamingXml(elementToXml(root), scope);
        expect(reread.status).toBe("parsed");
        if (reread.status !== "parsed") return;
        const rereadRoot = reread.value.elements?.at(0) ?? {};
        expectNamespacesResolved(rereadRoot, scope?.bindings ?? new Map());
        expect(elementToXml(rereadRoot)).toBe(elementToXml(root));
      }),
      propertyConfig({ numRuns: 300 }),
    );
  });
});

test.each(["' injected='yes", '"/>', "&quot;", "1.5", "-1"])(
  "refuses a non-decimal attribute replacement: %s",
  (replacement) => {
    expect(
      rewriteStreamingXmlDecimalAttributes(
        '<r id="1"/>',
        () => new Map([["id", replacement]]),
      ),
    ).toEqual({ status: "unsupported" });
  },
);
