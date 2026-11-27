import { expect, test } from "bun:test";
import { readFileSync } from "node:fs";
import path from "node:path";

import { localeMessages } from "./localeMessages";
import { FOLIO_LOCALES } from "./messages/locales";

const catalog = (locale: string): unknown =>
  JSON.parse(
    readFileSync(
      path.join(import.meta.dir, "messages", `${locale}.json`),
      "utf8",
    ),
  );

test.each(FOLIO_LOCALES.filter((locale) => locale !== "en"))(
  "the %s catalog built from its strings is the catalog its JSON holds",
  async (locale) => {
    const { messages } = (await import(`./messages/${locale}.gen`)) as {
      messages: unknown;
    };
    expect(messages).toEqual(catalog(locale));
  },
);

test("a catalog with a string too few or too many is refused", () => {
  expect(() => localeMessages([])).toThrow();
  const english = catalog("en");
  const count = (value: unknown): number =>
    typeof value === "string"
      ? 1
      : Object.values(value as Record<string, unknown>).reduce<number>(
          (total, child) => total + count(child),
          0,
        );
  const strings = Array.from({ length: count(english) + 1 }, () => "x");
  expect(() => localeMessages(strings)).toThrow();
});
