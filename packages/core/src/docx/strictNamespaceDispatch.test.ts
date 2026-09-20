/**
 * A Strict package's children reach the same disposition as a Transitional one's.
 *
 * ISO Strict spells every OOXML namespace under `purl.oclc.org`, and the
 * dispatcher's namespace-keyed dispositions are written in Transitional. A
 * lookup on the URI as written sends a Strict document's maths to the verbatim
 * sink: the equation survives as bytes and stops being an equation, so nothing
 * renders it, edits it or reads its text.
 *
 * The first test runs over the generated namespace table rather than over
 * maths alone, because the defect belongs to the lookup and not to one
 * namespace: the next disposition keyed by a URI would inherit it.
 */

import { describe, expect, test } from "bun:test";

import { CAPTURE, dispatchChildren } from "./containerChildren";
import { CONTAINER_CHILDREN } from "./containerChildren.gen";
import { parseParagraph } from "./paragraphParser";
import { TRANSITIONAL_NAMESPACE_BY_STRICT_URI } from "./strictValueEncodings.gen";
import { getChildElements, parseXml, parseXmlDocument, type XmlElement } from "./xmlParser";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const STRICT_W = "http://purl.oclc.org/ooxml/wordprocessingml/main";
const STRICT_MATH = "http://purl.oclc.org/ooxml/officeDocument/math";

const DECLARED = CONTAINER_CHILDREN["w:comment"];

// SAFETY: the entries come from `DECLARED` itself, so the record is total over
// its element type; `Object.fromEntries` only widens the key back to `string`.
const handlers = Object.fromEntries(DECLARED.map((name) => [name, CAPTURE] as const)) as Record<
  (typeof DECLARED)[number],
  typeof CAPTURE
>;

/** The `w:comment` inside a wrapper, so the fixture's prefixes are in scope. */
const container = (xml: string): XmlElement => {
  const root = parseXml(xml);
  return getChildElements(root).at(0) ?? root;
};

/**
 * Every pair but WordprocessingML's own.
 *
 * A child in the main namespace, Strict or Transitional, is a *declared*
 * child: `WORDPROCESSINGML_NAMESPACE_URIS` already carries both spellings and
 * the handler map decides it by local name, so it never reaches a
 * namespace-keyed disposition.
 */
const FOREIGN_NAMESPACE_PAIRS = [...TRANSITIONAL_NAMESPACE_BY_STRICT_URI].filter(
  ([strict]) => strict !== STRICT_W,
);

describe("the child dispatcher under ISO Strict", () => {
  test.each(FOREIGN_NAMESPACE_PAIRS)(
    "routes a child in %s to the disposition its Transitional pair carries",
    (strict, transitional) => {
      const reached: string[] = [];
      const preserved = dispatchChildren({
        element: container(`<w:comment xmlns:w="${W}" xmlns:s="${strict}"><s:thing/></w:comment>`),
        container: "w:comment",
        handlers,
        capturePosition: () => 0,
        undeclaredNamespaces: {
          [transitional]: (child) => {
            reached.push(child.name ?? "");
          },
        },
      });

      expect(reached).toEqual(["s:thing"]);
      expect(preserved).toBeUndefined();
    },
  );

  test("reads maths in a Strict paragraph as an equation, not as captured markup", () => {
    const root = parseXmlDocument(
      `<w:p xmlns:w="${STRICT_W}" xmlns:m="${STRICT_MATH}">` +
        `<m:oMath><m:r><m:t>x</m:t></m:r></m:oMath>` +
        `</w:p>`,
    ) as XmlElement | null;
    if (!root) {
      throw new Error("fixture did not parse");
    }
    const paragraph = parseParagraph(root, null, null, null, null, null);

    expect(paragraph.content.map((content) => content.type)).toEqual(["mathEquation"]);
  });
});
