/**
 * The run-level arm of the verbatim sink.
 *
 * `parseRunContents` recognises the run children folio models and used to let
 * the rest fall off the end of its switch. That is not merely a loss: the keep
 * rule in `paragraphParser` asks the *source* element whether a run carried a
 * payload while `serializeRun` writes the *model*, so an unmodelled child makes
 * the two disagree for exactly one save — save 1 emits a run with nothing in
 * it, the next parse drops that run, and save 2 differs from save 1. Holding
 * the markup turns the whole class into a fixed point at once.
 *
 * A preserved child is opaque everywhere except text: `w:ruby` puts its
 * `w:rubyBase` on the line as the word a reader reads, so the sink carries the
 * visible text alongside the markup and text extraction, markdown and layout
 * keep working.
 */

import type { PreservedInline, PreservedXmlContent } from "../types/document";

import { captureVerbatimXml } from "./verbatimCapture";
import {
  findChildrenByLocalName,
  getChildElements,
  getLocalName,
  getTextContent,
  type XmlElement,
} from "./xmlParser";

/**
 * Where a preserved run child hides text a reader sees, by local name: the
 * child elements whose `w:t` descendants are on the line rather than above,
 * beside, or nowhere.
 *
 * `w:ruby` is the whole list today. Its `w:rt` is the annotation printed above
 * the base and is not the sentence's text; its `w:rubyBase` is.
 */
const VISIBLE_TEXT_CHILDREN: ReadonlyMap<string, readonly string[]> = new Map([
  ["ruby", ["rubyBase"]],
]);

/** Cap on the text one preserved child contributes, mirroring `w:t` handling. */
const MAX_PRESERVED_TEXT_LENGTH = 100_000;

const collectTextContent = (element: XmlElement, into: string[]): void => {
  for (const child of getChildElements(element)) {
    if (getLocalName(child.name) === "t") {
      into.push(getTextContent(child));
      continue;
    }
    collectTextContent(child, into);
  }
};

/**
 * The visible text a preserved run child contributes, or `""` when it shows
 * nothing. Unknown markup contributes nothing on purpose: guessing at the text
 * of an element folio has never seen would put invented words in a document.
 */
export const preservedRunChildText = (element: XmlElement): string => {
  const sources = VISIBLE_TEXT_CHILDREN.get(getLocalName(element.name));
  if (sources === undefined) {
    return "";
  }
  const parts: string[] = [];
  for (const localName of sources) {
    for (const source of findChildrenByLocalName(element, localName)) {
      collectTextContent(source, parts);
    }
  }
  return parts.join("").slice(0, MAX_PRESERVED_TEXT_LENGTH);
};

/**
 * Capture one unmodelled run child.
 *
 * `captureVerbatimXml` materialises exactly the namespace bindings the
 * fragment's own prefixes need and leaves the canonical ones to the rebuilt
 * root, so ordinary `w:` markup comes back byte for byte while a foreign
 * prefix still arrives bound. Copying the whole root scope onto the fragment
 * instead would rewrite every capture with declarations it does not use.
 */
export const preserveRunChild = (element: XmlElement): PreservedXmlContent => ({
  type: "preservedXml",
  xml: captureVerbatimXml(element),
  text: preservedRunChildText(element),
});

/**
 * Where an unmodelled *inline* child hides text a reader sees, by local name.
 *
 * `w:customXml` and `w:smartTag` are transparent wrappers (ECMA-376 §17.5.1,
 * §17.5.1.9): their content is ordinary inline content, so their `w:t`
 * descendants are on the line. A `w:hyperlink` and a `w:fldSimple` are here
 * for the same reason and only in one place — inside another link or another
 * field, where folio captures the inner one rather than modelling the nesting.
 * Their runs print like any other; a capture that kept the markup and not the
 * words would take a linked clause off the page.
 *
 * Every other inline child folio captures — `w:permStart`, `w:proofErr` and
 * the custom-XML revision ranges — is an empty marker, and an element folio
 * has never seen contributes nothing on purpose, because guessing at its text
 * would put invented words in a document.
 */
const VISIBLE_TEXT_INLINE_CHILDREN: ReadonlySet<string> = new Set([
  "customXml",
  "fldSimple",
  "hyperlink",
  "smartTag",
]);

/**
 * One capture from the shared dispatcher's sink, as a member of the
 * container's own inline union.
 *
 * The sink holds markup the schema does not declare for the container at all,
 * so nothing is known about it beyond its bytes: `text` is empty because a
 * name folio has never seen shows nothing folio can read.
 */
export const preservedInlineCapture = (xml: string): PreservedInline => ({
  type: "preservedInline",
  xml,
  text: "",
});

/** Capture one unmodelled inline child of a paragraph or an inline wrapper. */
export const preserveInlineChild = (element: XmlElement): PreservedInline => {
  const parts: string[] = [];
  if (VISIBLE_TEXT_INLINE_CHILDREN.has(getLocalName(element.name))) {
    collectTextContent(element, parts);
  }
  return {
    type: "preservedInline",
    xml: captureVerbatimXml(element),
    text: parts.join("").slice(0, MAX_PRESERVED_TEXT_LENGTH),
  };
};
