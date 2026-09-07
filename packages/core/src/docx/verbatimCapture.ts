/**
 * The one way to turn parsed markup into a string a rebuilt part replays.
 *
 * folio parses Strict and Transitional packages and rebuilds every package as
 * Transitional. Modeled properties are re-serialized, so they come out
 * Transitional whatever the source was; content the parser keeps verbatim —
 * raw property XML, drawing and shape bodies, text-box markup, unmodeled
 * extensions — does not, and a Strict fragment copied under a Transitional root
 * changes meaning: the namespaces it declares stay Strict while the root is
 * not, and its values are spelled the way Strict spells them (`155.85pt`,
 * `20%`) where the Transitional attribute is a number.
 *
 * A fragment is therefore converted here, at capture, where its own namespace
 * scope still says which conformance class produced it. Nothing downstream has
 * to know: {@link captureVerbatimXml} is the only producer of replayed markup,
 * and a fragment that carries no Strict namespace passes through byte for byte.
 */

import { TaggedError } from "better-result";

import {
  type PercentUnit,
  type SlotEncoding,
  TRANSITIONAL_NAMESPACE_BY_STRICT_URI,
} from "./strictValueEncodings.gen";
import {
  isStrictNamespaceUri,
  STRICT_URI_PREFIX,
  toTransitionalNamespaceUri,
  transitionalSlotEncoding,
} from "./transitionalSpelling";
import { roundHalfAwayFromZero, universalMeasureAs } from "./universalMeasure";
import { elementToXml, getLocalName, type XmlElement, type XmlNamespaceScope } from "./xmlParser";

/** A Strict fragment uses a namespace that has no Transitional counterpart. */
export class UntranslatableStrictNamespaceError extends TaggedError(
  "UntranslatableStrictNamespaceError",
)<{
  message: string;
  uri: string;
}> {}

const TRANSITIONAL_URIS: ReadonlySet<string> = new Set(
  TRANSITIONAL_NAMESPACE_BY_STRICT_URI.values(),
);

/**
 * Drawing extensions the repository has no schema for.
 *
 * `wp14:pctWidth` and its siblings are Microsoft vocabularies, but they are
 * DrawingML by construction and DrawingML has one integer encoding for a
 * percentage, so a Strict producer's `20%` there converts like `a:alpha`'s.
 * Lengths are left alone: only an attribute's own type says what unit its
 * number counts, and guessing one would corrupt the value silently.
 */
const DRAWING_EXTENSION_MARKERS: readonly string[] = [
  "/wordprocessingCanvas",
  "/wordprocessingDrawing",
  "/wordprocessingGroup",
  "/wordprocessingInk",
  "/wordprocessingShape",
];

const isUnschemadDrawingNamespace = (uri: string): boolean =>
  !TRANSITIONAL_URIS.has(uri) && DRAWING_EXTENSION_MARKERS.some((marker) => uri.includes(marker));

/** `-?12.5%`: the one shape ECMA-376 gives a percentage that carries its unit. */
const PERCENTAGE = /^(-?[0-9]+(?:\.[0-9]+)?)%$/u;

const NUMBERS_PER_PERCENT: Readonly<Record<PercentUnit, number>> = {
  fiftiethPercent: 50,
  thousandthPercent: 1000,
  wholePercent: 1,
};

/** The Transitional spelling of one Strict-produced value, or the value unchanged. */
const transitionalValue = (
  value: string,
  namespaceUri: string,
  encoding: SlotEncoding | undefined,
): string => {
  const measureUnit = encoding?.measure;
  if (measureUnit !== undefined) {
    const measure = universalMeasureAs(value, measureUnit);
    if (measure !== undefined) {
      return String(measure);
    }
  }

  const percentage = PERCENTAGE.exec(value);
  if (percentage === null) {
    return value;
  }
  const percentUnit =
    encoding?.percent ??
    (isUnschemadDrawingNamespace(namespaceUri) ? "thousandthPercent" : undefined);
  if (percentUnit === undefined) {
    return value;
  }
  // SAFETY: the capture group is present whenever the pattern matched.
  return String(roundHalfAwayFromZero(Number(percentage[1]!) * NUMBERS_PER_PERCENT[percentUnit]));
};

/**
 * Which conformance class produced a subtree's markup.
 *
 * Inherited rather than read per element: an extension such as
 * `wp14:sizeRelH` has one namespace in both classes, so only its OOXML
 * ancestry says how its producer spelled the values inside it.
 */
type FragmentOrigin = "strict" | "transitional";

const originOf = (element: XmlElement, inherited: FragmentOrigin): FragmentOrigin => {
  const uri = element.namespaceUri;
  if (uri === undefined) {
    return inherited;
  }
  if (isStrictNamespaceUri(uri)) {
    return "strict";
  }
  if (uri.startsWith(STRICT_URI_PREFIX)) {
    throw new UntranslatableStrictNamespaceError({
      message: `Strict namespace ${uri} has no Transitional counterpart; folio will not replay it under a Transitional root.`,
      uri,
    });
  }
  return TRANSITIONAL_URIS.has(uri) ? "transitional" : inherited;
};

const convertAttributes = (
  element: XmlElement,
  namespaceUri: string,
  localName: string,
): Record<string, string | number | undefined> | undefined => {
  const attributes = element.attributes ?? {};
  let converted: Record<string, string | number | undefined> | undefined;
  for (const [name, raw] of Object.entries(attributes)) {
    if (typeof raw !== "string") {
      continue;
    }
    const next =
      name === "xmlns" || name.startsWith("xmlns:")
        ? toTransitionalNamespaceUri(raw)
        : transitionalValue(
            raw,
            namespaceUri,
            transitionalSlotEncoding(element.namespaceUri, localName, getLocalName(name)),
          );
    if (next === raw) {
      continue;
    }
    converted ??= { ...attributes };
    converted[name] = next;
  }
  return converted;
};

const childWithValue = (child: XmlElement, text: string): XmlElement =>
  text === child.text ? child : { ...child, text };

const toTransitional = (element: XmlElement, inherited: FragmentOrigin): XmlElement => {
  if (element.type === "text") {
    return element;
  }

  const origin = originOf(element, inherited);
  const namespaceUri =
    element.namespaceUri === undefined ? "" : toTransitionalNamespaceUri(element.namespaceUri);
  const localName = getLocalName(element.name);
  const textEncoding = transitionalSlotEncoding(element.namespaceUri, localName);

  const attributes =
    origin === "strict" && element.attributes
      ? convertAttributes(element, namespaceUri, localName)
      : undefined;

  let elements: XmlElement[] | undefined;
  for (const [index, child] of (element.elements ?? []).entries()) {
    const next =
      origin === "strict" && child.type === "text" && typeof child.text === "string"
        ? childWithValue(child, transitionalValue(child.text, namespaceUri, textEncoding))
        : toTransitional(child, origin);
    if (next === child) {
      continue;
    }
    elements ??= [...(element.elements ?? [])];
    elements[index] = next;
  }

  if (attributes === undefined && elements === undefined) {
    return element;
  }
  return {
    ...element,
    ...(attributes === undefined ? {} : { attributes }),
    ...(elements === undefined ? {} : { elements }),
  };
};

/** True when any binding in scope, including inherited ones, is a Strict namespace. */
const scopeDeclaresStrict = (scope: XmlNamespaceScope | undefined): boolean => {
  for (let current = scope; current; current = current.parent) {
    for (const uri of current.bindings.values()) {
      if (uri.startsWith(STRICT_URI_PREFIX)) {
        return true;
      }
    }
  }
  return false;
};

/**
 * Serialize parsed markup for replay inside a rebuilt Transitional part.
 *
 * Every verbatim capture goes through here — a lint rule keeps
 * {@link elementToXml} out of reach of the parsers — so a replay path added
 * later cannot skip the conversion. Markup that binds no Strict namespace,
 * anywhere in scope or in the fragment itself, is serialized once and returned
 * unchanged.
 */
export const captureVerbatimXml = (element: XmlElement): string => {
  const xml = elementToXml(element);
  if (!xml.includes(STRICT_URI_PREFIX) && !scopeDeclaresStrict(element.namespaceScope)) {
    return xml;
  }
  return elementToXml(toTransitional(element, "transitional"));
};
