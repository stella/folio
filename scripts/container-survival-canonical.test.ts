/**
 * The law sees folio's canonical output form, and only where an entry says so.
 *
 * `CANONICAL_SPELLINGS` is the one hand-written table in the survival law, and
 * a wrong entry fails in the worst direction: a real loss reads as a survival.
 * So it is held from three sides — every entry is exercised by a pair the
 * census generates, every entry's citation still reads what it claims, and the
 * law keeps reporting a loss for the neighbouring spelling folio really does
 * drop.
 */

import { describe, expect, test } from "bun:test";

import { allSubjects } from "./container-survival-census";
import {
  CANONICAL_SPELLINGS,
  type CanonicalSpelling,
  type CanonicalSubject,
  canonicalSpellingsFor,
} from "./lib/container-survival/canonicalSpellings";
import { buildFixture, type Subject } from "./lib/container-survival/fixture";
import { forcedSavePart, runSurvivalLaws, subjectKey } from "./lib/container-survival/laws";
import { loadContainerSpace, WML_NAMESPACE } from "./lib/container-survival/schemaSpace";
import { valuesForType } from "./lib/container-survival/values";

const space = await loadContainerSpace();
const subjects = allSubjects(space);

type AttributeSubject = Extract<Subject, { kind: "attribute" }>;

/** The census's key with its one namespace spelled the way a document spells it. */
const readableKey = (subject: Subject): string =>
  subjectKey(subject).replaceAll(`{${WML_NAMESPACE}}`, "w:");

const subjectAt = (key: string): Subject => {
  const found = subjects.find((subject) => readableKey(subject) === key);
  if (found === undefined) {
    throw new Error(`no pair in the census space is keyed ${key}`);
  }
  return found;
};

const attributeSubjectAt = (key: string): AttributeSubject => {
  const found = subjectAt(key);
  if (found.kind !== "attribute") {
    throw new Error(`${key} is a child pair`);
  }
  return found;
};

const canonicalSubjectOf = (subject: Subject): CanonicalSubject => ({
  element: subject.kind === "child" ? subject.slot.child : subject.slot.container.element,
  type: subject.kind === "child" ? subject.slot.childTypeQName : subject.slot.container.typeQName,
  attribute: subject.kind === "attribute" ? subject.slot.attribute : undefined,
  value: subject.kind === "attribute" ? subject.value : undefined,
});

const describeEntry = (entry: CanonicalSpelling): string =>
  entry.kind === "value"
    ? `${entry.type}@${entry.attribute}=${entry.authored.join("|")}`
    : `${entry.type} ${entry.authored} -> ${entry.canonical}`;

/** The attribute pairs one entry applies to, with the value the census writes. */
const slotsFor = (entry: CanonicalSpelling): AttributeSubject[] =>
  subjects.filter(
    (subject): subject is AttributeSubject =>
      subject.kind === "attribute" &&
      canonicalSpellingsFor(canonicalSubjectOf(subject)).includes(entry),
  );

describe("every canonical spelling is one a pair exercises", () => {
  test("no entry sits in the table unmatched", () => {
    const matched = new Set<CanonicalSpelling>();
    for (const subject of subjects) {
      for (const entry of canonicalSpellingsFor(canonicalSubjectOf(subject))) {
        matched.add(entry);
      }
    }
    const unmatched = CANONICAL_SPELLINGS.filter((entry) => !matched.has(entry)).map(describeEntry);
    expect(unmatched).toEqual([]);
  });

  test("a value entry lists only spellings the attribute's own type accepts", () => {
    const wrong: string[] = [];
    for (const entry of CANONICAL_SPELLINGS) {
      if (entry.kind !== "value") {
        continue;
      }
      const slot = slotsFor(entry).at(0);
      if (slot === undefined) {
        wrong.push(`${describeEntry(entry)}: no attribute pair carries it`);
        continue;
      }
      const { values } = valuesForType(space.index, slot.slot.typeQName);
      const unknown = entry.authored.filter((value) => !values.includes(value));
      if (unknown.length > 0) {
        wrong.push(
          `${describeEntry(entry)}: ${unknown.join(",")} is outside ${slot.slot.typeQName}`,
        );
      }
    }
    expect(wrong).toEqual([]);
  });
});

describe("every canonical spelling cites the line that performs it", () => {
  test("each cited line still reads what its entry claims", async () => {
    const stale: string[] = [];
    for (const entry of CANONICAL_SPELLINGS) {
      for (const { file, line, writes } of entry.writtenBy) {
        // oxlint-disable-next-line no-await-in-loop -- nine small files, read once each
        const source = await Bun.file(new URL(`../${file}`, import.meta.url)).text();
        const cited = source.split("\n")[line - 1] ?? "";
        if (!cited.includes(writes)) {
          stale.push(`${describeEntry(entry)}: ${file}:${line} reads ${cited.trim()}`);
        }
      }
    }
    expect(stale).toEqual([]);
  });
});

const mechanismOf = async (subject: Subject): Promise<string> => {
  const outcome = await runSurvivalLaws(space, subject);
  return outcome.unrepresentable ?? outcome.mechanism ?? "survives";
};

describe("the law reads a rename and still reads a drop", () => {
  test("a logical-direction border and indent survive; an indent in character units does not", async () => {
    expect(await mechanismOf(subjectAt("w:tblBorders|w:CT_TblBorders/w:start"))).toBe("survives");
    expect(await mechanismOf(subjectAt("w:ind|w:CT_Ind@w:start"))).toBe("survives");
    expect(await mechanismOf(subjectAt("w:ind|w:CT_Ind@w:endChars"))).toBe(
      "serialized-only-via-verbatim-replay",
    );
  }, 180_000);
});

describe("an absent w:val is the on state and nothing else", () => {
  const keepNext = attributeSubjectAt("w:keepNext|w:CT_OnOff@w:val");

  test("the on state survives as the bare element", async () => {
    expect(keepNext.value).toBe("true");
    expect(await mechanismOf(keepNext)).toBe("survives");
  }, 60_000);

  test("the off state survives by keeping its value, not by the entry", async () => {
    const off: AttributeSubject = { kind: "attribute", slot: keepNext.slot, value: "0" };
    expect(await mechanismOf(off)).toBe("survives");
    const built = buildFixture(space, off);
    if (built.status !== "built") {
      throw new Error(built.reason);
    }
    expect(await forcedSavePart(built.fixture)).toContain('<w:keepNext w:val="0"/>');
    const onOff = CANONICAL_SPELLINGS.filter(
      (entry) => entry.kind === "value" && entry.type === "CT_OnOff",
    ).map((entry) => (entry.kind === "value" ? entry.authored.join("|") : entry.authored));
    expect(onOff).toEqual(["true|1|on"]);
  }, 60_000);
});
