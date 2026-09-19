/**
 * `CT_NonVisualDrawingProps` (`wp:docPr`, `wps:cNvPr`, `pic:cNvPr`) carries the
 * three authored strings every drawing object has: its name, its alt text
 * (`descr`) and its title. They are authored content, not folio's to mint, so
 * one reader and one writer own them for every drawing kind.
 */

import { escapeXmlAttribute } from "@stll/docx-core";
import type { XmlElement } from "./xmlParser";
import { getAttribute } from "./xmlParser";

/** The authored name, alt text and title of a drawing object. */
export type NonVisualDrawingNames = {
  /** `@name`, schema-required, so `""` means the object was never named. */
  name?: string;
  /** `@descr`: alt text, accessibility content. */
  alt?: string;
  /** `@title`. */
  title?: string;
};

/**
 * Read `@name`, `@descr` and `@title` off a `CT_NonVisualDrawingProps` element.
 *
 * An attribute that is not there stays absent. `@name` is schema-required, so
 * an unnamed object still writes one, and `name=""` is the marker the writer
 * below mints for exactly that: reading it back as an authored name would make
 * `absent → save → parse` land on `""` instead of absent, and no reader can
 * tell the two apart. `@descr` and `@title` are optional and written only when
 * authored, so `""` in either is a string someone wrote.
 */
export const parseNonVisualDrawingNames = (
  element: XmlElement | null | undefined,
): NonVisualDrawingNames => {
  if (!element) {
    return {};
  }
  const name = getAttribute(element, null, "name");
  const descr = getAttribute(element, null, "descr");
  const title = getAttribute(element, null, "title");
  return {
    ...(name !== null && name !== "" ? { name } : {}),
    ...(descr !== null ? { alt: descr } : {}),
    ...(title !== null ? { title } : {}),
  };
};

/**
 * The `name`, `descr` and `title` attributes of a `CT_NonVisualDrawingProps`
 * element, leading space included.
 *
 * `@name` is schema-required, so an object carrying none writes the empty
 * string. folio does not name an object it did not author: a generated
 * "Shape 3" would overwrite the author's own name for every drawing whose
 * name did not survive the model, and read back as authored content.
 */
export const serializeNonVisualDrawingNames = (names: NonVisualDrawingNames): string => {
  const descr = names.alt === undefined ? "" : ` descr="${escapeXmlAttribute(names.alt)}"`;
  const title = names.title === undefined ? "" : ` title="${escapeXmlAttribute(names.title)}"`;
  return ` name="${escapeXmlAttribute(names.name ?? "")}"${descr}${title}`;
};
