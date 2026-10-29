import { describe, expect, test } from "bun:test";

import { readXml, wellFormedXmlIssue } from "./streamingXmlParser";
import { FOLIO_XML_RESOURCE_LIMITS } from "./xmlResourceLimits";

const DECLARATION = '<?xml version="1.0" encoding="UTF-8"?>';

describe("wellFormedXmlIssue", () => {
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
