/**
 * The law sees folio's canonical output form, and only where an entry says so.
 *
 * A wrong entry in `CANONICAL_SPELLINGS` fails in the worst direction: a real
 * loss reads as a survival. Its renames are generated from the schema, so the
 * risk sits in what stays hand-written — the three omissions, and the decision
 * about where folio writes each generated rename. So the table is held from
 * three sides: every entry is exercised by a pair the census generates, every
 * entry's citation still reads what it claims, and the law keeps reporting a
 * loss for the neighbouring spelling folio really does drop.
 */

import { afterAll, describe, expect, test } from "bun:test";
import {
  isIdentifier,
  isVariableStatement,
  type Node,
  type SourceFile,
  SyntaxKind,
} from "tsrs/unstable/ast";
import { API } from "tsrs/unstable/async";

import { allSubjects } from "./container-survival-census";
import {
  CANONICAL_SPELLINGS,
  type CanonicalSpelling,
  canonicalSpellingsFor,
  type CanonicalSubject,
} from "./lib/container-survival/canonicalSpellings";
import {
  buildFixture,
  modelledCompanionFor,
  type Subject,
} from "./lib/container-survival/fixture";
import {
  forcedSavePart,
  runSurvivalLaws,
  subjectKey,
} from "./lib/container-survival/laws";
import {
  loadContainerSpace,
  WML_NAMESPACE,
} from "./lib/container-survival/schemaSpace";
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
  element:
    subject.kind === "child"
      ? subject.slot.child
      : subject.slot.container.element,
  type:
    subject.kind === "child"
      ? subject.slot.childTypeQName
      : subject.slot.container.typeQName,
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
    const unmatched = CANONICAL_SPELLINGS.filter(
      (entry) => !matched.has(entry),
    ).map(describeEntry);
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

// Citations are code fragments, not whole programs: they are tokenized with
// the TypeScript parser's error recovery rather than parsed strictly.
let compilerApi: API | undefined;
afterAll(async () => {
  await compilerApi?.close();
});

const withSourceFile = async <T>(
  fileName: string,
  source: string,
  read: (parsed: SourceFile) => T,
): Promise<T> => {
  compilerApi ??= new API({ cwd: import.meta.dir });
  const retained = await compilerApi.createSourceFile(fileName, source);
  try {
    return read(retained.sourceFile);
  } finally {
    await retained.dispose();
  }
};

const codeTokens = (source: string): Promise<string> =>
  withSourceFile("citation.ts", source, (parsed) => {
    const tokens: string[] = [];
    const visit = (node: Node): void => {
      const children = node.getChildren(parsed);
      if (children.length > 0) {
        for (const child of children) visit(child);
        return;
      }
      if (
        node.kind !== SyntaxKind.EndOfFileToken &&
        node.getWidth(parsed) > 0
      ) {
        tokens.push(node.getText(parsed));
      }
    };
    visit(parsed);
    return tokens.join("\0");
  });

const citationMatches = async (
  source: string,
  { declaration, writes }: { declaration: string; writes: string },
): Promise<boolean> => {
  const ownerText = await withSourceFile("serializer.ts", source, (parsed) => {
    const declarations = parsed.statements.flatMap((statement) =>
      isVariableStatement(statement)
        ? [...statement.declarationList.declarations]
        : [],
    );
    return declarations
      .find((node) => isIdentifier(node.name) && node.name.text === declaration)
      ?.getText(parsed);
  });
  if (ownerText === undefined) return false;
  return `\0${await codeTokens(ownerText)}\0`.includes(
    `\0${await codeTokens(writes)}\0`,
  );
};

describe("every canonical spelling cites the declaration that performs it", () => {
  test("citations survive formatting and reject a different writer or commented-out write", async () => {
    const citation = {
      declaration: "INDENT_LEFT",
      writes: 'TRANSITIONAL_NAME_BY_STRICT_NAME["CT_Ind @start"]',
    };
    for (const prefix of [
      "",
      "\n\n",
      "// An unrelated declaration follows.\n",
    ]) {
      for (const spacing of ["", " ", "\n  ", "/* spacing */"]) {
        const write = `TRANSITIONAL_NAME_BY_STRICT_NAME${spacing}[${spacing}"CT_Ind @start"${spacing}]`;
        expect(
          await citationMatches(
            `${prefix}const INDENT_LEFT = ${write};`,
            citation,
          ),
        ).toBe(true);
        expect(
          await citationMatches(
            `${prefix}const INDENT_RIGHT = ${write};`,
            citation,
          ),
        ).toBe(false);
        expect(
          await citationMatches(
            `${prefix}const INDENT_LEFT = "left"; // ${write}`,
            citation,
          ),
        ).toBe(false);
        expect(
          await citationMatches(
            `${prefix}const INDENT_LEFT = /* ${write} */ "left";`,
            citation,
          ),
        ).toBe(false);
      }
    }
    expect(
      await citationMatches(
        'const INDENT_LEFT = TRANSITIONAL_NAME_BY_STRICT_NAME["CT_Ind @end"];',
        citation,
      ),
    ).toBe(false);
    const templateCitation = {
      declaration: "serializeOnOffElement",
      writes: "return value ? `<w:${name}/>`",
    };
    expect(
      await citationMatches(
        'const serializeOnOffElement = (value, name) => { return value ? `<w:${name}/>` : `<w:${name} w:val="0"/>`; };',
        templateCitation,
      ),
    ).toBe(true);
    expect(
      await citationMatches(
        'const serializeOnOffElement = (value, name) => { return value ? `<w:${name} w:val="0"/>` : `<w:${name}/>`; };',
        templateCitation,
      ),
    ).toBe(false);
  });

  test("each cited declaration still reads what its entry claims", async () => {
    const stale: string[] = [];
    for (const entry of CANONICAL_SPELLINGS) {
      for (const { file, declaration, writes } of entry.writtenBy) {
        // oxlint-disable-next-line no-await-in-loop -- nine small files, read once each
        const source = await Bun.file(
          new URL(`../${file}`, import.meta.url),
        ).text();
        if (!(await citationMatches(source, { declaration, writes }))) {
          stale.push(
            `${describeEntry(entry)}: ${file}:${declaration} does not write ${writes}`,
          );
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

describe("the law reads canonical and preserved spellings", () => {
  test("logical-direction borders and indents survive, including preserved character units", async () => {
    expect(
      await mechanismOf(subjectAt("w:tblBorders|w:CT_TblBorders/w:start")),
    ).toBe("survives");
    expect(await mechanismOf(subjectAt("w:ind|w:CT_Ind@w:start"))).toBe(
      "survives",
    );
    expect(await mechanismOf(subjectAt("w:ind|w:CT_Ind@w:endChars"))).toBe(
      "survives",
    );
  }, 180_000);

  test("a companion comes from another model field, not an alternate spelling", () => {
    const start = attributeSubjectAt("w:ind|w:CT_Ind@w:start");

    expect(modelledCompanionFor(space, start.slot)?.spelled).toBe("w:right");
  });
});

describe("an absent w:val is the on state and nothing else", () => {
  const keepNext = attributeSubjectAt("w:keepNext|w:CT_OnOff@w:val");

  test("the on state survives as the bare element", async () => {
    expect(keepNext.value).toBe("true");
    expect(await mechanismOf(keepNext)).toBe("survives");
  }, 60_000);

  test("the off state survives by keeping its value, not by the entry", async () => {
    const off: AttributeSubject = {
      kind: "attribute",
      slot: keepNext.slot,
      value: "0",
    };
    expect(await mechanismOf(off)).toBe("survives");
    const built = buildFixture(space, off);
    if (built.status !== "built") {
      throw new Error(built.reason);
    }
    expect(await forcedSavePart(built.fixture)).toContain(
      '<w:keepNext w:val="0"/>',
    );
    const onOff = CANONICAL_SPELLINGS.filter(
      (entry) => entry.kind === "value" && entry.type === "CT_OnOff",
    ).map((entry) =>
      entry.kind === "value" ? entry.authored.join("|") : entry.authored,
    );
    expect(onOff).toEqual(["true|1|on"]);
  }, 60_000);
});
