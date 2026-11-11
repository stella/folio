import { expect, test } from "bun:test";

import en from "../i18n/messages/en.json";
import { FOLIO_LOCALES } from "../i18n/messages/locales";
import {
  CANONICAL_GAP,
  CANONICAL_GAP_MESSAGE_KEY,
  CANONICAL_RULE,
  CANONICAL_RULE_MESSAGE_KEYS,
  canonicalRefusalNotice,
} from "./canonicalCapabilities";

const lookup = (messages: unknown, key: string): unknown => {
  let node: unknown = (messages as { folio: unknown }).folio;
  for (const part of key.split(".")) {
    if (typeof node !== "object" || node === null) return undefined;
    node = (node as Record<string, unknown>)[part];
  }
  return node;
};

test("a refusal under any gap shows the catalog's message, not the session's", () => {
  for (const gap of Object.values(CANONICAL_GAP)) {
    const asked: string[] = [];
    expect(
      canonicalRefusalNotice({ gap }, (key) => {
        asked.push(key);
        return String(lookup(en, key));
      }),
    ).toBe("This change can't be made here yet.");
    expect(asked).toEqual([CANONICAL_GAP_MESSAGE_KEY]);
  }
});

test("a rule refusal shows the editor's own message for the rule", () => {
  const asked: string[] = [];
  expect(
    canonicalRefusalNotice(
      { rule: CANONICAL_RULE.contentControlLock },
      (key) => {
        asked.push(key);
        return String(lookup(en, key));
      },
    ),
  ).toBe("This part of the document is locked.");
  expect(asked).toEqual(["editor.contentLocked"]);
});

test("every refusal message exists in every locale", async () => {
  for (const locale of FOLIO_LOCALES) {
    // oxlint-disable-next-line no-await-in-loop -- one catalog at a time
    const messages: unknown = (await import(`../i18n/messages/${locale}.json`))
      .default;
    for (const key of [
      ...Object.values(CANONICAL_RULE_MESSAGE_KEYS),
      CANONICAL_GAP_MESSAGE_KEY,
    ]) {
      const text = lookup(messages, key);
      expect([locale, key, typeof text]).toEqual([locale, key, "string"]);
    }
  }
});
