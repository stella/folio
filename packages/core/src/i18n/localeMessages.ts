import { panic } from "better-result";

import type { FolioMessages } from "./messages";
import { messages as english } from "./messages/en.gen";

/**
 * A locale's catalog from its strings alone, in the order the English
 * catalog holds its keys: every catalog has the English catalog's shape, so
 * the shape ships once and each locale carries only its strings.
 */
export const localeMessages = (values: readonly string[]): FolioMessages => {
  let next = 0;
  const build = (shape: Readonly<Record<string, unknown>>) => {
    const built: Record<string, unknown> = {};
    for (const key of Object.keys(shape)) {
      const value = shape[key];
      if (typeof value === "string") {
        built[key] = values[next] ?? panic("A locale catalog lacks a string.");
        next += 1;
      } else {
        built[key] = build(value as Readonly<Record<string, unknown>>);
      }
    }
    return built;
  };
  const messages = build(english);
  if (next !== values.length)
    panic("A locale catalog holds more strings than the English one.");
  // SAFETY: built from the English catalog's shape, one string a key.
  return messages as FolioMessages;
};
