/**
 * No reader hand-lists a Strict spelling the generated table already names.
 *
 * `packages/core/src/docx/strictNames.gen.ts` says which WordprocessingML names
 * ECMA-376 Part 1 spells by writing direction, and `strictNames.ts` is the one
 * place that takes both spellings. A second `?? findChild(…, "start")` copied
 * into a parser is the drift this guard exists to stop: it reads correctly on
 * the day it is written and silently stops matching the table the survival law
 * measures against the moment either side moves.
 *
 * The rule is the pair, not the name. `w:lvl/w:start` is a list's first number
 * and `w:lnNumType/@w:start` is a line-numbering origin — neither is a rename,
 * and neither has a Transitional partner to read beside it. So a finding is a
 * reader that takes **both** spellings of one pair off the same receiver, which
 * is exactly the shape the tolerance had before it moved into the table. There
 * are no exemptions: a finding is fixed by calling the table's readers.
 *
 * Usage:
 *   bun scripts/check-strict-name-tolerance.ts
 */

import path from "node:path";

import { TaggedError } from "better-result";

import {
  STRICT_NAMES_BY_TRANSITIONAL_NAME,
  TRANSITIONAL_NAME_BY_STRICT_NAME,
} from "../packages/core/src/docx/strictNames.gen";

const REPO_ROOT = path.resolve(import.meta.dir, "..");

class StrictNameToleranceError extends TaggedError("StrictNameToleranceError")<{
  message: string;
}> {}

/** The local name on each side of a rename, without the type the table keys by. */
const localNameOf = (key: string): string => {
  const local = key.slice(key.indexOf(" ") + 1);
  return local.startsWith("@") ? local.slice(1) : local;
};

/**
 * Each renamed name to the names on the other side of its pair.
 *
 * A set rather than one name: nothing says a Transitional name stands for a
 * single Strict spelling, and a map that kept only the last would drop a
 * hand-listed reader without saying so.
 */
const partnersByName = (): ReadonlyMap<string, ReadonlySet<string>> => {
  const partners = new Map<string, Set<string>>();
  const pair = (name: string, partner: string): void => {
    partners.set(name, (partners.get(name) ?? new Set()).add(partner));
  };
  for (const [key, transitional] of Object.entries(TRANSITIONAL_NAME_BY_STRICT_NAME)) {
    pair(localNameOf(key), transitional);
  }
  for (const [key, spellings] of Object.entries(STRICT_NAMES_BY_TRANSITIONAL_NAME)) {
    for (const strict of spellings) {
      pair(localNameOf(key), strict);
    }
  }
  return partners;
};

/**
 * A WordprocessingML read: the receiver, and the name it reads off it.
 *
 * The receiver is matched as an identifier path, so a read off the result of
 * another call does not match. That is the trade the guard makes: it never
 * reports a reader that is not hand-listing a spelling, and the idiom it looks
 * for — two reads of one element, side by side — always names its element.
 */
const READER_CALL =
  /\b(?:findChild|findChildren|findChildByLocalName|getAttribute|getAttributeAnyPrefix|parseNumericAttribute|parseOnOffAttribute|parseTableMeasurementValue)\(\s*([A-Za-z_$][\w$]*(?:[.?]?\.?\w+)*)\s*,\s*(?:"w"|'w'|null)\s*,\s*["']([A-Za-z0-9]+)["']/gu;

/** Every place one source reads both spellings of a rename off one element. */
export const handListedSpellings = (source: string, file = "<source>"): string[] => {
  const partners = partnersByName();
  const readsByReceiver = new Map<string, Map<string, number>>();
  for (const match of source.matchAll(READER_CALL)) {
    const [, receiver, name] = match;
    if (receiver === undefined || name === undefined || !partners.has(name)) {
      continue;
    }
    const reads = readsByReceiver.get(receiver) ?? new Map<string, number>();
    if (!reads.has(name)) {
      reads.set(name, source.slice(0, match.index).split("\n").length);
    }
    readsByReceiver.set(receiver, reads);
  }

  const findings: string[] = [];
  for (const [receiver, reads] of readsByReceiver) {
    for (const [name, line] of reads) {
      for (const partner of partners.get(name) ?? []) {
        if (reads.has(partner) && name < partner) {
          findings.push(
            `${file}:${line} reads w:${name} and w:${partner} off ${receiver}; the rename belongs to strictNames.gen.ts, so read it with findChildAnySpelling/numericAttributeAnySpelling/hasAttributeAnySpelling.`,
          );
        }
      }
    }
  }
  return findings;
};

const main = async (): Promise<void> => {
  const sources = new Bun.Glob("packages/*/src/**/*.ts").scan({ cwd: REPO_ROOT });
  const findings: string[] = [];
  for await (const file of sources) {
    if (file.endsWith(".test.ts")) {
      continue;
    }
    findings.push(...handListedSpellings(await Bun.file(path.join(REPO_ROOT, file)).text(), file));
  }
  if (findings.length > 0) {
    throw new StrictNameToleranceError({ message: findings.toSorted().join("\n") });
  }
  process.stdout.write("no reader hand-lists a Strict spelling\n");
};

if (import.meta.main) {
  await main();
}
