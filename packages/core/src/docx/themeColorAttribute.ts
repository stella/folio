/**
 * The one reader for a theme-colour attribute.
 *
 * `w:themeColor`, `w:themeFill` and the `w:clrSchemeMapping` values are all
 * `ST_ThemeColor`. Before this the attribute was narrowed against a union that
 * carried DrawingML slot names instead, so `hyperlink` — the token Word writes
 * for a hyperlink's colour — failed to narrow and the attribute was dropped at
 * parse time, taking it out of the next save. The enumeration is now generated
 * from the schema, and a token outside it is captured rather than refused, the
 * way `CT_Border/@w:val` already keeps a style outside `ST_Border`.
 */

import { PARSE_WARNING_CODES, readThemeColor, type ThemeColorValue } from "@stll/docx-core/model";

import type { ParseContext } from "./parseContext";

export type ThemeColorAttributeOptions = {
  /** The attribute as written, or `null` when the element does not carry it. */
  raw: string | null | undefined;
  /** The element the attribute sits on, for the warning's location. */
  element?: string | undefined;
  context?: ParseContext | undefined;
};

export const parseThemeColorAttribute = ({
  raw,
  element,
  context,
}: ThemeColorAttributeOptions): ThemeColorValue | undefined => {
  const value = readThemeColor(raw);
  if (value !== undefined && typeof value !== "string") {
    context?.warn({
      code: PARSE_WARNING_CODES.unrecognisedThemeColor,
      value: value.raw,
      ...(element === undefined ? {} : { element }),
    });
  }
  return value;
};
