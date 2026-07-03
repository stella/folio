#!/usr/bin/env bun
/**
 * Port translations from archived eigenpal/docx-editor-i18n locale JSON into
 * folio's `packages/react/src/i18n/messages/*.json` shape by matching identical
 * English source strings.
 *
 * Usage: bun scripts/port-upstream-folio-locales.ts <upstream-locale-json> <folio-out-json>
 *
 * Example:
 *   bun scripts/port-upstream-folio-locales.ts /tmp/docx-editor-upstream/packages/i18n/he.json packages/react/src/i18n/messages/he.json
 */

import { readFileSync, writeFileSync } from "node:fs";

type Nested = { [key: string]: string | Nested };

const flatten = (obj: Nested, prefix = ""): Map<string, string> => {
  const out = new Map<string, string>();
  for (const [key, value] of Object.entries(obj)) {
    const path = prefix ? `${prefix}.${key}` : key;
    if (value === null || value === undefined) {
      continue;
    }
    if (typeof value === "string") {
      out.set(path, value);
    } else {
      for (const [childKey, childValue] of flatten(value, path)) {
        out.set(childKey, childValue);
      }
    }
  }
  return out;
};

const setNested = (obj: Nested, path: string, value: string): void => {
  const parts = path.split(".");
  let current: Nested = obj;
  for (const part of parts.slice(0, -1)) {
    const next = current[part];
    if (typeof next === "object" && next !== null) {
      current = next;
      continue;
    }
    const child: Nested = {};
    current[part] = child;
    current = child;
  }
  const leaf = parts.at(-1);
  if (leaf) {
    current[leaf] = value;
  }
};

if (!upstreamLocalePath || !folioOutPath) {
  console.error(
    "Usage: bun scripts/port-upstream-folio-locales.ts <upstream-locale.json> <folio-out.json>",
  );
  process.exit(1);
}

const folioEnPath = "packages/react/src/i18n/messages/en.json";
const upstreamEnPath = "/tmp/docx-editor-upstream/packages/i18n/en.json";

const folioEn = JSON.parse(readFileSync(folioEnPath, "utf8")) as { folio: Nested };
const upstreamEn = JSON.parse(readFileSync(upstreamEnPath, "utf8")) as Nested;
const upstreamLocale = JSON.parse(readFileSync(upstreamLocalePath, "utf8")) as Nested;

const folioFlat = flatten(folioEn.folio, "folio");
const upstreamEnFlat = flatten(upstreamEn);
const upstreamLocaleFlat = flatten(upstreamLocale);

// English value -> upstream key (first wins on collision)
const enValueToUpstreamKey = new Map<string, string>();
for (const [key, value] of upstreamEnFlat) {
  if (!enValueToUpstreamKey.has(value)) {
    enValueToUpstreamKey.set(value, key);
  }
}

const out: Nested = { folio: {} };
let matched = 0;
let fallback = 0;

for (const [folioKey, english] of folioFlat) {
  const upstreamKey = enValueToUpstreamKey.get(english);
  const translated =
    upstreamKey !== undefined ? upstreamLocaleFlat.get(upstreamKey) : undefined;
  if (translated !== undefined && translated !== english) {
    setNested(out, folioKey, translated);
    matched++;
  } else if (translated !== undefined) {
    setNested(out, folioKey, translated);
    matched++;
  } else {
    setNested(out, folioKey, english);
    fallback++;
  }
}

writeFileSync(folioOutPath, `${JSON.stringify(out, null, 2)}\n`);
console.error(`Wrote ${folioOutPath}: ${matched} matched, ${fallback} English fallback`);
