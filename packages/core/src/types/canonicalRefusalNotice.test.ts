import { expect, test } from "bun:test";

import en from "../i18n/messages/en.json";
import { FOLIO_LOCALES } from "../i18n/messages/locales";
import {
  CANONICAL_GAP,
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

test("a gap refusal shows the session's message", () => {
  expect(
    canonicalRefusalNotice(
      { gap: CANONICAL_GAP.trackedHyperlinkResolution },
      "Hyperlink suggestions are unavailable.",
      () => "translated",
    ),
  ).toBe("Hyperlink suggestions are unavailable.");
});

test("a rule refusal shows the editor's own message for the rule", () => {
  const asked: string[] = [];
  expect(
    canonicalRefusalNotice(
      { rule: CANONICAL_RULE.contentControlLock },
      "The content of a locked content control cannot change.",
      (key) => {
        asked.push(key);
        return String(lookup(en, key));
      },
    ),
  ).toBe("This part of the document is locked.");
  expect(asked).toEqual(["editor.contentLocked"]);
});

test("every rule's message exists in every locale", async () => {
  for (const locale of FOLIO_LOCALES) {
    // oxlint-disable-next-line no-await-in-loop -- one catalog at a time
    const messages: unknown = (await import(`../i18n/messages/${locale}.json`))
      .default;
    for (const key of Object.values(CANONICAL_RULE_MESSAGE_KEYS)) {
      const text = lookup(messages, key);
      expect([locale, key, typeof text]).toEqual([locale, key, "string"]);
    }
  }
});
