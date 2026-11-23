import { describe, expect, spyOn, test } from "bun:test";

import { readXml, wellFormedXmlIssue } from "./streamingXmlParser";
import { FOLIO_XML_RESOURCE_LIMITS } from "./xmlResourceLimits";

const DECLARATION = '<?xml version="1.0" encoding="UTF-8"?>';

describe("wellFormedXmlIssue", () => {
  test("checks in-scope namespace prefixes when requested", () => {
    const options = { namespaces: "require-bound-prefixes" } as const;
    expect(
      wellFormedXmlIssue(
        '<root xmlns:p="urn:outer"><p:child xmlns:q="urn:inner" q:item="1" xml:space="preserve"/><p:next/></root>',
        options,
      ),
    ).toBeNull();
    expect(
      wellFormedXmlIssue('<p:root xmlns:p="urn:root" p:attr="1"/>', options),
    ).toBeNull();
    expect(wellFormedXmlIssue("<root><p:child/></root>")).toBeNull();
    expect(wellFormedXmlIssue("<root><p:child/></root>", options)).toBe(
      "an unbound namespace prefix p",
    );
    expect(
      wellFormedXmlIssue(
        '<root xmlns="urn:default"><child xmlns:p="urn:local"/><next p:item="1"/></root>',
        options,
      ),
    ).toBe("an unbound namespace prefix p");
  });

  test("admits only XML whitespace beside the root", () => {
    expect(wellFormedXmlIssue(`${DECLARATION}\n<a/>\r\n\t `)).toBeNull();
    expect(wellFormedXmlIssue(`\ufeff${DECLARATION}<a/>`)).toBeNull();
    for (const space of [" ", " ", "﻿"]) {
      expect(wellFormedXmlIssue(`<a/>${space}`)).not.toBeNull();
    }
  });

  test("refuses CDATA outside the root element, even when it holds nothing", () => {
    expect(wellFormedXmlIssue("<a><![CDATA[ ]]></a>")).toBeNull();
    expect(wellFormedXmlIssue("<![CDATA[ ]]><a/>")).not.toBeNull();
    expect(wellFormedXmlIssue("<a/><![CDATA[]]>")).not.toBeNull();
  });

  test("requires the declaration to carry a version", () => {
    expect(wellFormedXmlIssue("<?xml version='1.0'?><a/>")).toBeNull();
    expect(
      wellFormedXmlIssue('<?xml version="1.1" standalone="yes"?><a/>'),
    ).toBeNull();
    for (const declaration of [
      "<?xml?>",
      '<?xml encoding="UTF-8"?>',
      '<?xml version="2.0"?>',
      "<?xml version=\"1.0'?>",
      '<?xml versionx="1.0"?>',
    ]) {
      expect(wellFormedXmlIssue(`${declaration}<a/>`)).not.toBeNull();
    }
  });
});

describe("the nesting limit", () => {
  const nested = (depth: number, innermost: string): string =>
    `${"<a>".repeat(depth - 1)}${innermost}${"</a>".repeat(depth - 1)}`;
  const { maxDepth } = FOLIO_XML_RESOURCE_LIMITS;

  test("counts a self-closing element like an open one", () => {
    for (const innermost of ["<b/>", "<b></b>"]) {
      expect(
        readXml(nested(maxDepth, innermost), { conformance: "ooxml" }).status,
      ).toBe("parsed");
      expect(
        readXml(nested(maxDepth + 1, innermost), { conformance: "ooxml" })
          .status,
      ).toBe("rejected");
    }
  });
});

test("quoted attribute width does not add JavaScript character probes", () => {
  const read = (width: number) => {
    const left = 'text>"'.repeat(width);
    const right = "text>'".repeat(width);
    const xml = `<p left='${left}' right="${right}"><t/></p>`;
    const probes = spyOn(String.prototype, "charCodeAt");
    try {
      const parsed = readXml(xml, { conformance: "ooxml" });
      const work = probes.mock.calls.length;
      expect(parsed.status).toBe("parsed");
      if (parsed.status === "parsed") {
        const root = parsed.value.elements?.at(0);
        expect(root?.attributes).toEqual({ left, right });
        expect(root?.elements?.at(0)?.name).toBe("t");
      }
      return work;
    } finally {
      probes.mockRestore();
    }
  };
  expect(read(1024)).toBe(read(1));
  for (const xml of ['<p a="unclosed>text</p>', "<p a='unclosed>text</p>"]) {
    expect(readXml(xml, { conformance: "ooxml" }).status).toBe("rejected");
  }
});
