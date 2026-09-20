/**
 * However a revision and a transparent wrapper are nested, a save keeps the
 * nesting it was given.
 *
 * `w:ins`/`w:del`/`w:moveFrom`/`w:moveTo` and `w:bdo`/`w:dir`/`w:sdt` compose
 * freely and in either order: each says something the other does not, so a
 * save that reorders them changes what the document claims. The example tests
 * cover one level; the input class is every stack up to three deep, which is
 * where a segmentation pass that treats one of the two as the segment owner
 * shows itself.
 */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";

import { parseParagraph } from "./paragraphParser";
import { serializeParagraph } from "./serializer/paragraphSerializer";
import { parseXmlDocument, type XmlElement } from "./xmlParser";

setDefaultTimeout(propertyTestTimeout(30_000));

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const ATTRS = 'w:id="3" w:author="A" w:date="2026-01-01T00:00:00Z"';

type Layer = { open: string; close: string; removes: boolean | null };

const REVISIONS: Layer[] = [
  { open: `<w:ins ${ATTRS}>`, close: "</w:ins>", removes: false },
  { open: `<w:del ${ATTRS}>`, close: "</w:del>", removes: true },
  { open: `<w:moveFrom ${ATTRS}>`, close: "</w:moveFrom>", removes: true },
  { open: `<w:moveTo ${ATTRS}>`, close: "</w:moveTo>", removes: false },
];

/** A transparent wrapper says nothing about removal, hence `null`. */
const WRAPPERS: Layer[] = [
  { open: '<w:bdo w:val="rtl">', close: "</w:bdo>", removes: null },
  { open: '<w:dir w:val="ltr">', close: "</w:dir>", removes: null },
  { open: "<w:sdt><w:sdtPr/><w:sdtContent>", close: "</w:sdtContent></w:sdt>", removes: null },
];

/**
 * Write the stack around a run.
 *
 * The innermost revision decides how the run's text is spelled: `w:del` and
 * `w:moveFrom` hold `w:delText`, the other two hold `w:t`, and a transparent
 * wrapper between the revision and the run changes neither. A revision nested
 * inside another is the inner one's answer, which is the spelling folio
 * writes for `<w:del><w:ins>`.
 */
const markup = (layers: readonly Layer[]): string => {
  let removed = false;
  for (const layer of layers) {
    removed = layer.removes ?? removed;
  }

  let stack = removed ? "<w:r><w:delText>x</w:delText></w:r>" : "<w:r><w:t>x</w:t></w:r>";
  for (let index = layers.length - 1; index >= 0; index -= 1) {
    // SAFETY: index walks the array's own bounds downwards.
    const layer = layers[index]!;
    stack = `${layer.open}${stack}${layer.close}`;
  }
  return stack;
};

const save = (inner: string): string => {
  const node = parseXmlDocument(`<w:p ${NS}>${inner}</w:p>`) as XmlElement | null;
  if (!node) {
    throw new Error("the fixture did not parse");
  }
  return serializeParagraph(parseParagraph(node, new Map(), null, null));
};

const bodyOf = (paragraphXml: string): string =>
  paragraphXml.replace(/^<w:p>/u, "").replace(/<\/w:p>$/u, "");

const layerStack = fc.array(fc.oneof(...[...REVISIONS, ...WRAPPERS].map(fc.constant)), {
  minLength: 1,
  maxLength: 3,
});

describe("nesting a revision and a transparent wrapper", () => {
  test("a save keeps the authored stack", () => {
    fc.assert(
      fc.property(layerStack, (layers) => {
        const authored = markup(layers);
        expect(bodyOf(save(authored))).toBe(authored);
      }),
      propertyConfig(),
    );
  });

  test("a second save changes nothing", () => {
    fc.assert(
      fc.property(layerStack, (layers) => {
        const once = bodyOf(save(markup(layers)));
        expect(bodyOf(save(once))).toBe(once);
      }),
      propertyConfig(),
    );
  });
});
