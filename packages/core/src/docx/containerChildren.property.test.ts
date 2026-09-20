/**
 * The dispatcher keeps every child it did not model, byte-equal and in place.
 *
 * The generator interleaves declared children (which the container models),
 * undeclared ones (a foreign namespace, an `mc:` construct, an element a later
 * OOXML revision adds), because interleaving is the part a
 * hand-rolled `switch` gets wrong: a `default` that appends to the end of the
 * container reorders the content model, and one that does nothing at all
 * drops it. Both pass a test that only counts.
 *
 * The round trip is the real thing — capture, serialize, parse the result —
 * rather than a comparison of two in-memory records, so a capture that cannot
 * be replayed under a rebuilt root fails here.
 */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";

import { CAPTURE, dispatchChildren, serializeWithPreservedChildren } from "./containerChildren";
import { CONTAINER_CHILDREN } from "./containerChildren.gen";
import { getChildElements, getLocalName, parseXml, type XmlElement } from "./xmlParser";

setDefaultTimeout(propertyTestTimeout(30_000));

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const VENDOR = "urn:example:vendor";

/** Markup for one generated child, and what the container should do with it. */
type GeneratedChild =
  | { declared: true; name: string; xml: string }
  | { declared: false; xml: string };

const DECLARED = CONTAINER_CHILDREN["w:comment"];

/**
 * Undeclared children a real package writes: a foreign namespace, an `mc:`
 * construct, and an element from a later revision of the same namespace.
 */
const undeclaredChild = fc.constantFrom(
  '<x:note xmlns:x="urn:example:vendor" x:kind="aside">kept</x:note>',
  '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">' +
    "<mc:Fallback><w:p/></mc:Fallback></mc:AlternateContent>",
  '<w:futureElement w:val="7"/>',
);

const generatedChild: fc.Arbitrary<GeneratedChild> = fc.oneof(
  fc
    .constantFrom(...DECLARED)
    .map((name): GeneratedChild => ({ declared: true, name, xml: `<w:${name}/>` })),
  undeclaredChild.map((xml): GeneratedChild => ({ declared: false, xml })),
);

/**
 * A container folio models `w:p` in and captures everything else, which is the
 * shape every migrated container has: one or two modelled children, a sink for
 * the rest.
 */
// SAFETY: the entries come from `DECLARED` itself, so the record is total
// over its element type; `Object.fromEntries` only widens the key back to
// `string`.
const handlers = Object.fromEntries(DECLARED.map((name) => [name, CAPTURE] as const)) as Record<
  (typeof DECLARED)[number],
  typeof CAPTURE
>;

const roundTrip = (
  children: readonly GeneratedChild[],
): { xml: string; modelledNames: string[] } => {
  const source = parseXml(
    `<w:comment xmlns:w="${W}" xmlns:x="${VENDOR}" w:id="1">` +
      `${children.map(({ xml }) => xml).join("")}</w:comment>`,
  );
  const element = getChildElements(source).at(0) ?? source;

  const modelled: string[] = [];
  const preserved = dispatchChildren({
    element,
    container: "w:comment",
    capturePosition: () => modelled.length,
    handlers: {
      ...handlers,
      p: (child: XmlElement) => modelled.push(`<w:${getLocalName(child.name)}/>`),
    },
  });

  return {
    xml: `<w:comment w:id="1">${serializeWithPreservedChildren(modelled, preserved)}</w:comment>`,
    modelledNames: modelled,
  };
};

/** The top element of a fragment, by local name. */
const fragmentName = (xml: string): string => {
  const root = parseXml(`<root xmlns:w="${W}" xmlns:x="${VENDOR}">${xml}</root>`);
  const fragment = getChildElements(getChildElements(root).at(0) ?? root).at(0);
  return getLocalName(fragment?.name);
};

/** The container's children, by local name, in document order. */
const childNames = (xml: string): string[] => {
  const root = parseXml(`<root xmlns:w="${W}" xmlns:x="${VENDOR}">${xml}</root>`);
  const container = getChildElements(getChildElements(root).at(0) ?? root).at(0);
  return getChildElements(container).map((child) => getLocalName(child.name));
};

describe("the shared child dispatcher", () => {
  test("keeps every undeclared child, in position, and models the declared ones", () => {
    fc.assert(
      fc.property(fc.array(generatedChild, { maxLength: 8 }), (children) => {
        const { xml, modelledNames } = roundTrip(children);

        // Position: the container's child sequence is the source's, by local
        // name, whether folio modelled a child or only kept it.
        expect(childNames(xml)).toEqual(
          children.map((child) => (child.declared ? child.name : fragmentName(child.xml))),
        );

        // Modelled: only `w:p` reaches the model; nothing else pretends to.
        expect(modelledNames).toEqual(
          children.filter((child) => child.declared && child.name === "p").map(() => "<w:p/>"),
        );

        // Byte-equal: every undeclared child's markup survives verbatim.
        for (const child of children) {
          if (!child.declared) {
            expect(xml).toContain(child.xml);
          }
        }
      }),
      propertyConfig({ numRuns: 200 }),
    );
  });

  test("a foreign-namespace child never reaches the handler its local name matches", () => {
    const element = getChildElements(
      parseXml(
        `<root xmlns:w="${W}" xmlns:x="${VENDOR}">` +
          `<w:comment><x:p x:kind="aside">kept</x:p><w:p/></w:comment></root>`,
      ),
    )
      .at(0)
      ?.elements?.find((node): node is XmlElement => node.type === "element");
    const modelled: string[] = [];
    const preserved = dispatchChildren({
      element: element ?? parseXml("<w:comment/>"),
      container: "w:comment",
      capturePosition: () => modelled.length,
      handlers: { ...handlers, p: () => modelled.push("<w:p/>") },
    });

    // `x:p` is not the `w:p` the content model declares. Modelling it would
    // read a foreign element as a paragraph and write back whatever folio's
    // paragraph serializer makes of it, which is the source's markup gone.
    expect(modelled).toEqual(["<w:p/>"]);
    expect(serializeWithPreservedChildren(modelled, preserved)).toContain('x:kind="aside"');
  });

  test("a container with nothing unmodelled carries no sink at all", () => {
    const element = getChildElements(
      parseXml(`<root xmlns:w="${W}"><w:comment><w:p/><w:p/></w:comment></root>`),
    )
      .at(0)
      ?.elements?.find((node): node is XmlElement => node.type === "element");
    const modelled: string[] = [];
    expect(
      dispatchChildren({
        element: element ?? parseXml("<w:comment/>"),
        container: "w:comment",
        capturePosition: () => modelled.length,
        handlers: { ...handlers, p: () => modelled.push("<w:p/>") },
      }),
    ).toBeUndefined();
  });
});
