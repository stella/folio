import { describe, expect, test } from "bun:test";

import type { Document } from "../types/document";
import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import { parseDocumentBody } from "./documentParser";
import { serializeDocument } from "./serializer/documentSerializer";

const TRANSITIONAL_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const STRICT_NAMESPACE = "http://purl.oclc.org/ooxml/wordprocessingml/main";

const documentXml = (prefix: string, namespace: string): string =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<${prefix}:document xmlns:${prefix}="${namespace}">` +
  `<${prefix}:background ${prefix}:color="auto" ${prefix}:themeColor="none" ` +
  `${prefix}:themeTint="66" ${prefix}:themeShade="BF">` +
  `<${prefix}:drawing/>` +
  `</${prefix}:background>` +
  `<${prefix}:body><${prefix}:p><${prefix}:r><${prefix}:t>Body</${prefix}:t>` +
  `</${prefix}:r></${prefix}:p></${prefix}:body></${prefix}:document>`;

const editorRoundTrip = (source: string): string => {
  const document: Document = { package: { document: parseDocumentBody(source) } };
  return serializeDocument(fromProseDoc(toProseDoc(document), document));
};

describe("document background round trip", () => {
  for (const profile of [
    { name: "Transitional", prefix: "w", namespace: TRANSITIONAL_NAMESPACE },
    { name: "Strict with an alternate prefix", prefix: "s", namespace: STRICT_NAMESPACE },
  ]) {
    test(`${profile.name} background is an editor-save fixed point`, () => {
      const first = editorRoundTrip(documentXml(profile.prefix, profile.namespace));
      const second = editorRoundTrip(first);

      expect(second).toBe(first);
      expect(first).toContain(
        '<w:background w:color="auto" w:themeColor="none" w:themeTint="66" w:themeShade="BF">',
      );
      expect(first).toMatch(/<(?:w|s):drawing\b/u);
      expect(first.indexOf("<w:background")).toBeLessThan(first.indexOf("<w:body>"));
    });
  }

  test("a same-named foreign element is not treated as a document background", () => {
    const source =
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<w:document xmlns:w="${TRANSITIONAL_NAMESPACE}" xmlns:x="urn:foreign">` +
      `<x:background x:color="auto"/><w:body/></w:document>`;

    expect(parseDocumentBody(source).background).toBeUndefined();
  });
});
