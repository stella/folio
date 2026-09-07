/**
 * Keep the application version a package states about itself in the form the
 * schema gives it.
 *
 * `AppVersion` in the extended-properties part is `XX.YYYY`: a one- or
 * two-digit integer, a dot, and four digits. Producers exist that write a
 * three-part version there instead, and folio copies `docProps/app.xml`
 * through verbatim when it saves a document it did not create — so a package
 * can carry a value with two dots in, and a save that copies it out hands a
 * consumer a package it refuses to open at all.
 *
 * {@link appVersionInSchemaForm} is the one mapping, and it is a pure function
 * of the value alone: it keeps the leading integer where the value opens with
 * one the form allows, and writes the build digits the form requires. Nothing
 * else in the part is touched, and a package that carries no extended
 * properties keeps carrying none — synthesizing metadata a document never
 * stated is a different decision.
 */

import { TaggedError } from "better-result";

/** `XX.YYYY`: the only form the extended-properties `AppVersion` may take. */
const SCHEMA_FORM = /^\d{1,2}\.\d{4}$/u;

/** The leading integer, when the value opens with one the form allows. */
const LEADING_MAJOR = /^\d{1,2}(?!\d)/u;

/** Major version for a value that does not open with one. */
const DEFAULT_MAJOR = "1";

/**
 * The build digits a rewritten value gets. A value the form rejects states no
 * build this pass could carry over, so every rewrite lands on the same one.
 */
const CANONICAL_BUILD = "0000";

/**
 * A value reached the package that the schema form does not accept.
 *
 * The normalization hands every value it writes to one choke point, which
 * throws this rather than letting the package go out carrying a version a
 * consumer refuses.
 */
export class AppVersionSchemaError extends TaggedError("AppVersionSchemaError")<{
  message: string;
  appVersion: string;
}> {}

/**
 * `value` when it already fits, and a value derived from it when it does not.
 *
 * The derivation reads the value and nothing else — not the package, not the
 * clock, not folio's own version — so two saves of one document state the same
 * application version.
 */
export const appVersionInSchemaForm = (value: string): string => {
  if (SCHEMA_FORM.test(value)) {
    return value;
  }
  const major = LEADING_MAJOR.exec(value.trim())?.[0] ?? DEFAULT_MAJOR;
  const normalized = `${major}.${CANONICAL_BUILD}`;
  if (!SCHEMA_FORM.test(normalized)) {
    throw new AppVersionSchemaError({
      message: `Derived application version ${normalized} is not of the form XX.YYYY`,
      appVersion: value,
    });
  }
  return normalized;
};

/**
 * The element and its content, with the prefix and attributes it was written
 * with preserved: the replacement rewrites the value between the tags and
 * leaves the rest of the part byte-identical.
 */
const APP_VERSION_ELEMENT =
  /(<(?:[^\s<>/:]+:)?AppVersion(?:\s[^<>]*)?>)([^<]*)(<\/(?:[^\s<>/:]+:)?AppVersion>)/u;

/**
 * Rewrite the application version an extended-properties part states.
 *
 * Returns the part unchanged when the value already fits and when the part
 * states no version at all, so a save of a document that never carried a
 * malformed one is byte-identical.
 */
export const normalizeAppVersionInExtendedProperties = (xml: string): string =>
  xml.replace(APP_VERSION_ELEMENT, (whole, open: string, value: string, close: string) => {
    const replacement = appVersionInSchemaForm(value);
    return replacement === value ? whole : `${open}${replacement}${close}`;
  });
