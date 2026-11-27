import { CATALOG_SHAPE, type CatalogShape } from "./catalogShape.gen";
import type { FolioMessages } from "./messages";

/**
 * A locale's catalog from its strings alone, in the order the English
 * catalog holds its keys: every catalog has the English catalog's shape, so
 * the shape ships once and each locale carries only its strings. A locale
 * entry reads no other locale and no package.
 */
export const localeMessages = (values: readonly string[]): FolioMessages => {
  let next = 0;
  const build = (shape: CatalogShape) => {
    const built: Record<string, unknown> = {};
    for (const key of Object.keys(shape)) {
      const value = shape[key];
      if (value === 0) {
        const string = values[next];
        if (string === undefined)
          throw new Error("A locale catalog lacks a string.");
        built[key] = string;
        next += 1;
      } else if (value !== undefined) {
        built[key] = build(value);
      }
    }
    return built;
  };
  const messages = build(CATALOG_SHAPE);
  if (next !== values.length)
    throw new Error(
      "A locale catalog holds more strings than the English one.",
    );
  // SAFETY: built from the English catalog's shape, one string a key.
  return messages as FolioMessages;
};
