import { describe, expect, test } from "bun:test";
import fc from "fast-check";

import {
  buildForbiddenRules,
  emptyGlossary,
  findDroppedExactSelectors,
  findDroppedPlurals,
  findForbiddenTerms,
  findIcuError,
  findMissingPluralCategories,
  findPlaceholderMismatch,
  isSuppressed,
  parseGlossary,
} from "./i18n-lint";
import type { LintBaseline } from "./i18n-lint";

describe("findPlaceholderMismatch", () => {
  test("returns null when the variable sets match (order is irrelevant)", () => {
    expect(findPlaceholderMismatch("Hi {name}, {count}", "{count} – {name}")).toBeNull();
  });

  test("flags a dropped variable", () => {
    expect(findPlaceholderMismatch("Sent to {email}", "Odesláno")).toEqual({
      missing: ["email"],
      extra: [],
    });
  });

  test("flags a renamed variable as both missing and extra", () => {
    expect(findPlaceholderMismatch("{a}", "{b}")).toEqual({
      missing: ["a"],
      extra: ["b"],
    });
  });

  test("tracks rich-text tag names (t.rich), not just their children", () => {
    expect(findPlaceholderMismatch("<terms>Terms</terms>", "<link>Podmínky</link>")).toEqual({
      missing: ["<terms>"],
      extra: ["<link>"],
    });
  });

  test("counts the plural argument, not the # placeholder", () => {
    expect(
      findPlaceholderMismatch(
        "{count, plural, one {# item} other {# items}}",
        "{count, plural, one {# položka} few {# položky} other {# položek}}",
      ),
    ).toBeNull();
  });
});

describe("findIcuError", () => {
  test("returns null for valid ICU", () => {
    expect(findIcuError("{count, plural, one {#} other {#}}")).toBeNull();
  });

  test("returns a message for broken ICU", () => {
    expect(findIcuError("{count, plural, one {#}")).not.toBeNull();
  });
});

describe("findMissingPluralCategories", () => {
  test("flags Polish missing few/many for an English-shaped plural", () => {
    expect(findMissingPluralCategories("{n, plural, one {#} other {#}}", "pl")).toEqual(
      expect.arrayContaining(["n#few", "n#many"]),
    );
  });

  test("passes when the locale needs only the present categories", () => {
    expect(findMissingPluralCategories("{n, plural, one {#} other {#}}", "de")).toEqual([]);
  });

  test("does not require cs/sk fractional-only `many` for integer UI counts", () => {
    expect(findMissingPluralCategories("{n, plural, one {#} few {#} other {#}}", "cs")).toEqual([]);
  });

  test("does not require the dormant es/fr/pt-BR `many` (exact millions)", () => {
    for (const locale of ["es", "fr", "pt-BR"]) {
      expect(findMissingPluralCategories("{n, plural, one {#} other {#}}", locale)).toEqual([]);
    }
  });

  test("does not count exact selectors like =0 toward CLDR categories", () => {
    expect(findMissingPluralCategories("{n, plural, =0 {none} one {#} other {#}}", "en")).toEqual(
      [],
    );
  });

  test("degrades to [] for an invalid locale code instead of throwing", () => {
    expect(findMissingPluralCategories("{n, plural, one {#} other {#}}", "!")).toEqual([]);
  });
});

describe("findDroppedPlurals", () => {
  test("flags a source plural flattened to plain interpolation", () => {
    expect(findDroppedPlurals("{count, plural, one {#} other {#}}", "{count} items")).toEqual([
      "count",
    ]);
  });

  test("passes when the target keeps the plural", () => {
    expect(
      findDroppedPlurals(
        "{count, plural, one {#} other {#}}",
        "{count, plural, one {#} few {#} other {#}}",
      ),
    ).toEqual([]);
  });

  test("flags a plural that dropped the count from every branch", () => {
    expect(
      findDroppedPlurals(
        "{count, plural, one {# file} other {# files}}",
        "{count, plural, one {file} other {files}}",
      ),
    ).toEqual(["count"]);
  });

  test("allows one branch to omit the count when another shows it", () => {
    expect(
      findDroppedPlurals(
        "{count, plural, one {# file} other {# files}}",
        "{count, plural, one {jeden soubor} other {# souborů}}",
      ),
    ).toEqual([]);
  });

  test("does not count a nested plural's # as the outer count", () => {
    expect(
      findDroppedPlurals(
        "{count, plural, one {# file} other {# files}}",
        "{count, plural, one {{n, plural, other {#}}} other {{n, plural, other {#}}}}",
      ),
    ).toEqual(["count"]);
  });

  test("counts a {count, number} formatted node as showing the count", () => {
    expect(
      findDroppedPlurals(
        "{count, plural, one {# file} other {# files}}",
        "{count, plural, one {{count, number} soubor} other {{count, number} souborů}}",
      ),
    ).toEqual([]);
  });
});

describe("findDroppedExactSelectors", () => {
  test("flags a dropped source exact selector (=0 zero-state)", () => {
    expect(
      findDroppedExactSelectors(
        "{count, plural, =0 {No results} one {# result} other {# results}}",
        "{count, plural, one {# výsledek} few {# výsledky} other {# výsledků}}",
      ),
    ).toEqual(["count=0"]);
  });

  test("passes when the target keeps the source's exact selectors", () => {
    expect(
      findDroppedExactSelectors(
        "{count, plural, =0 {No results} other {# results}}",
        "{count, plural, =0 {Žádné} other {# výsledků}}",
      ),
    ).toEqual([]);
  });
});

describe("isSuppressed", () => {
  const baseline: LintBaseline = {
    placeholder: { "a.b": { cs: { source: "Open {x}", target: "Otevřít" } } },
    icu: {},
    plural: {},
    terminology: {},
  };

  test("suppresses only when both source and target are unchanged", () => {
    expect(
      isSuppressed(baseline, "placeholder", "a.b", "cs", {
        source: "Open {x}",
        target: "Otevřít",
      }),
    ).toBe(true);
  });

  test("re-checks when the translation is edited", () => {
    expect(
      isSuppressed(baseline, "placeholder", "a.b", "cs", {
        source: "Open {x}",
        target: "Edited",
      }),
    ).toBe(false);
  });

  test("re-checks when the en source changes under a stale translation", () => {
    expect(
      isSuppressed(baseline, "placeholder", "a.b", "cs", {
        source: "Open {x} and {y}",
        target: "Otevřít",
      }),
    ).toBe(false);
  });
});

describe("parseGlossary", () => {
  test("an empty glossary yields no rules (terminology check is a no-op)", () => {
    expect(buildForbiddenRules(emptyGlossary())).toEqual([]);
  });

  test("term groups are optional and default to empty", () => {
    expect(parseGlossary("{}")).toEqual(emptyGlossary());
  });

  test("rejects a malformed term group instead of silently disabling it", () => {
    expect(() => parseGlossary(JSON.stringify({ verbs: "nope" }))).toThrow();
  });

  test("rejects a term with a non-array forbidden entry", () => {
    expect(() =>
      parseGlossary(JSON.stringify({ nouns: [{ id: "x", en: "X", forbidden: { de: "Wort" } }] })),
    ).toThrow();
  });
});

describe("terminology", () => {
  const rules = buildForbiddenRules(
    parseGlossary(
      JSON.stringify({
        legalConcepts: [{ id: "matter", en: "Matter", forbidden: { de: ["Sache"] } }],
      }),
    ),
  );

  test("flags a forbidden rendering when the source is about the concept", () => {
    expect(findForbiddenTerms("Open this matter", "Diese Sache öffnen", "de", rules)).toEqual([
      "Sache",
    ]);
  });

  test("does not fire when the source is unrelated to the concept", () => {
    expect(findForbiddenTerms("A factual question", "Eine reine Sache", "de", rules)).toEqual([]);
  });

  test("matches whole words only (no substring false positives)", () => {
    expect(findForbiddenTerms("Open this matter", "Sachenrecht gilt", "de", rules)).toEqual([]);
  });

  test("detects the concept when the source uses the plural trigger", () => {
    // "matters" (plural) still triggers the "Matter" concept.
    expect(findForbiddenTerms("Manage your matters", "Diese Sache öffnen", "de", rules)).toEqual([
      "Sache",
    ]);
  });

  test("matches an Arabic forbidden term carrying proclitics (ال/و prefixes)", () => {
    const arRules = buildForbiddenRules(
      parseGlossary(
        JSON.stringify({
          legalConcepts: [{ id: "matter", en: "Matter", forbidden: { ar: ["قضية"] } }],
        }),
      ),
    );
    // bare, ال-prefixed, and و+ال-prefixed forms all count.
    expect(findForbiddenTerms("Open this matter", "افتح قضية", "ar", arRules)).toEqual(["قضية"]);
    expect(findForbiddenTerms("Open this matter", "افتح القضية", "ar", arRules)).toEqual(["قضية"]);
    expect(findForbiddenTerms("Open this matter", "والقضية مفتوحة", "ar", arRules)).toEqual([
      "قضية",
    ]);
  });
});

describe("terminology: key triggers and forbiddenOnKey", () => {
  const rules = buildForbiddenRules(
    parseGlossary(
      JSON.stringify({
        legalConcepts: [
          {
            id: "team",
            en: "Team",
            // firm: broad (unconditional). org: key-trigger always, word-trigger
            // unless the source is about "organiz-ing".
            forbidden: { de: ["Kanzlei"], en: ["firm"] },
            forbiddenOnKey: { de: ["Organisation"], en: ["organisation"] },
            keyTriggers: ["scopeTeam"],
            sourceExempt: ["organiz", "organis"],
          },
        ],
      }),
    ),
  );

  test("key trigger fires the rule even when English lacks the trigger word", () => {
    expect(
      findForbiddenTerms(
        "Everyone in the organisation",
        "Alle in der Organisation",
        "de",
        rules,
        "knowledge.skills.form.scopeTeam",
      ),
    ).toEqual(["Organisation"]);
  });

  test("forbiddenOnKey fires on the word trigger when the source is not exempt", () => {
    // Source says "team member" (no "organiz-"); a target rendering it as an
    // organization member is caught even on a non-scope key.
    expect(
      findForbiddenTerms(
        "Choose an existing team member",
        "Ein Organisation hinzufügen",
        "de",
        rules,
        "workspaces.members.addMemberDescription",
      ),
    ).toEqual(["Organisation"]);
  });

  test("sourceExempt suppresses word-trigger enforcement of forbiddenOnKey", () => {
    // Source also says "organize", so a Slavic/German org form renders that
    // word, not the team concept: not flagged on a non-scope key.
    expect(
      findForbiddenTerms(
        "Organize team activity",
        "Alle in der Organisation",
        "de",
        rules,
        "workspaces.emptyMatters.description",
      ),
    ).toEqual([]);
  });

  test("key trigger ignores sourceExempt (scope labels are always strict)", () => {
    expect(
      findForbiddenTerms(
        "Organize the team",
        "Alle in der Organisation",
        "de",
        rules,
        "knowledge.agentSkills.scopeTeam",
      ),
    ).toEqual(["Organisation"]);
  });

  test("forbidden (firm) words are enforced broadly via the word trigger", () => {
    expect(
      findForbiddenTerms(
        "team-wide list",
        "kanzleiweite Kanzlei",
        "de",
        rules,
        "settings.anonymization.description",
      ),
    ).toEqual(["Kanzlei"]);
  });

  test("English self-check flags banned source wording on a key-trigger key", () => {
    expect(
      findForbiddenTerms(
        "Everyone in the organisation",
        "Everyone in the organisation",
        "en",
        rules,
        "knowledge.agentSkills.scopeTeam",
      ),
    ).toEqual(["organisation"]);
  });

  test("no key and no word trigger means no enforcement", () => {
    expect(
      findForbiddenTerms(
        "Configure your organisation",
        "Alle in der Organisation",
        "de",
        rules,
        "consent.scopeOnboarding",
      ),
    ).toEqual([]);
  });
});

describe("terminology search semantics", () => {
  type MakeRulesOptions = {
    keyTriggers?: string[];
    sourceExempt?: string[];
  };
  const makeRules = (
    needle: string,
    { keyTriggers = [], sourceExempt = [] }: MakeRulesOptions = {},
  ) =>
    buildForbiddenRules(
      parseGlossary(
        JSON.stringify({
          nouns: [
            {
              id: "concept",
              en: "Matter",
              keyTriggers,
              sourceExempt,
              forbidden: { cs: [needle], ar: [needle], en: [needle] },
              forbiddenOnKey: { cs: ["restricted"] },
            },
          ],
        }),
      ),
    );

  test("retains letter-only boundaries, accents, exact Czech inflections and literal punctuation", () => {
    const rules = makeRules("Žaneta");
    for (const target of ["ŽANETA", "(Žaneta)", "1Žaneta_", "😀Žaneta", "Žaneta\u0301"]) {
      expect(findForbiddenTerms("Matters", target, "cs", rules)).toEqual(["Žaneta"]);
    }
    for (const target of ["Zaneta", "Žanety", "aŽaneta", "Žanetaa"]) {
      expect(findForbiddenTerms("Matters", target, "cs", rules)).toEqual([]);
    }
    expect(findForbiddenTerms("Matter", "a+b", "cs", makeRules("a+b"))).toEqual(["a+b"]);
    expect(findForbiddenTerms("Matter", "aaab", "cs", makeRules("a+b"))).toEqual([]);
  });

  test("keeps listed hit order and duplicates across overlapping terms", () => {
    const rules = buildForbiddenRules(
      parseGlossary(
        JSON.stringify({
          nouns: [
            {
              id: "concept",
              en: "Matter",
              forbidden: { cs: ["a-b", "b", "a-b"] },
            },
          ],
        }),
      ),
    );
    expect(findForbiddenTerms("Matter", "a-b a-b", "cs", rules)).toEqual(["a-b", "b", "a-b"]);
    expect(findForbiddenTerms("Matter", "anything", "cs", [])).toEqual([]);
  });

  test("preserves empty literal triggers, key case and source lowercase exemptions", () => {
    expect(
      findForbiddenTerms("Other", "banned", "cs", makeRules("banned", { keyTriggers: [""] }), ""),
    ).toEqual(["banned"]);
    expect(
      findForbiddenTerms(
        "Other",
        "banned",
        "cs",
        makeRules("banned", { keyTriggers: ["Scope"] }),
        "scope",
      ),
    ).toEqual([]);
    expect(
      findForbiddenTerms(
        "MATTERS ORGANIZE",
        "restricted",
        "cs",
        makeRules("banned", { sourceExempt: ["ORGAN"] }),
      ),
    ).toEqual([]);
    expect(
      findForbiddenTerms("Matter", "restricted", "cs", makeRules("banned", { sourceExempt: [""] })),
    ).toEqual([]);
    expect(findForbiddenTerms("Matter", "restricted", "cs", makeRules("banned"))).toEqual([
      "restricted",
    ]);
  });

  test("matches malformed patterns and text as normalized scalars while returning original terms", () => {
    for (const malformed of ["\ud800", "\udc00", "a\ud800-b", "a\udc00-b"]) {
      const normalized = malformed.toWellFormed();
      expect(findForbiddenTerms("Matter", normalized, "cs", makeRules(malformed))).toEqual([
        malformed,
      ]);
      expect(findForbiddenTerms("Matter", malformed, "cs", makeRules(normalized))).toEqual([
        normalized,
      ]);
      expect(
        findForbiddenTerms(
          "Other",
          "banned",
          "cs",
          makeRules("banned", { keyTriggers: [malformed] }),
          normalized,
        ),
      ).toEqual(["banned"]);
      expect(
        findForbiddenTerms(
          `Matter ${malformed}`,
          "restricted",
          "cs",
          makeRules("banned", { sourceExempt: [normalized] }),
        ),
      ).toEqual([]);
    }
    expect(
      findForbiddenTerms(
        "Other",
        "banned",
        "cs",
        makeRules("banned", { keyTriggers: ["\ud83d"] }),
        "😀",
      ),
    ).toEqual([]);
    expect(
      findForbiddenTerms("Other", "banned", "cs", makeRules("banned", { keyTriggers: [""] }), "😀"),
    ).toEqual(["banned"]);
    expect(
      findForbiddenTerms(
        "Matter 😀",
        "restricted",
        "cs",
        makeRules("banned", { sourceExempt: [""] }),
      ),
    ).toEqual([]);
  });

  test("matches substring-list oracles for configured key and source searches", () => {
    const fragment = fc.constantFrom(
      "",
      "a",
      "a+b",
      "Žaneta",
      "žaneta",
      "😀",
      ".scope",
      "_",
      "K",
      "\ud800",
      "\udc00",
      "a\ud800-b",
      "�",
    );
    fc.assert(
      fc.property(
        fc.array(fragment, { maxLength: 5 }),
        fc.array(fragment, { maxLength: 5 }),
        fragment,
        fragment,
        (keyTriggers, sourceExempt, key, sourceFragment) => {
          const source = `Matter ${sourceFragment}`;
          const keyFires = keyTriggers.some((trigger) =>
            key.toWellFormed().includes(trigger.toWellFormed()),
          );
          const isExempt = sourceExempt.some((exempt) =>
            source.toLowerCase().toWellFormed().includes(exempt.toLowerCase().toWellFormed()),
          );
          const expected = keyFires || !isExempt ? ["restricted"] : [];
          expect(
            findForbiddenTerms(
              source,
              "restricted",
              "cs",
              makeRules("banned", { keyTriggers, sourceExempt }),
              key,
            ),
          ).toEqual(expected);
        },
      ),
      { seed: 27018, numRuns: 300 },
    );
  });

  test("observes mutations to configured literal lists", () => {
    const keyTriggers = ["first"];
    const sourceExempt = ["organ"];
    const rules = makeRules("banned", { keyTriggers, sourceExempt });
    const rule = rules.at(0);
    expect(rule).toBeDefined();
    if (!rule) {
      return;
    }
    expect(findForbiddenTerms("Other", "banned", "cs", rules, "first")).toEqual(["banned"]);
    rule.keyTriggers.splice(0, 1, "second");
    expect(findForbiddenTerms("Other", "banned", "cs", rules, "first")).toEqual([]);
    expect(findForbiddenTerms("Other", "banned", "cs", rules, "second")).toEqual(["banned"]);
    expect(findForbiddenTerms("Matter organ", "restricted", "cs", rules)).toEqual([]);
    rule.sourceExempt.splice(0, 1, "else");
    expect(findForbiddenTerms("Matter organ", "restricted", "cs", rules)).toEqual(["restricted"]);
  });

  test("preserves Unicode case folding for long s and Kelvin sign", () => {
    expect(findForbiddenTerms("Matter", "ſ", "cs", makeRules("s"))).toEqual(["s"]);
    expect(findForbiddenTerms("Matter", "K", "cs", makeRules("k"))).toEqual(["k"]);
    expect(findForbiddenTerms("Matter", "S", "cs", makeRules("ſ"))).toEqual(["ſ"]);
    expect(findForbiddenTerms("Matter", "K", "cs", makeRules("K"))).toEqual(["K"]);
  });

  test("matches the normalized Unicode regex oracle over generated boundaries and locales", () => {
    const escape = (value: string) => value.replace(/[.*+?^${}()|[\]\\]/gu, "\\$&");
    fc.assert(
      fc.property(
        fc.constantFrom("cs", "ar", "en"),
        fc.constantFrom(
          "Žaneta",
          "قضية",
          "a+b",
          "Sache",
          "",
          "😀",
          "s",
          "k",
          "ſ",
          "K",
          "\ud800",
          "\udc00",
          "a\ud800-b",
          "�",
        ),
        fc.constantFrom(
          "",
          "a",
          "č",
          "1",
          "_",
          "😀",
          "\u0301",
          "(",
          "وال",
          "\ud800",
          "\udc00",
          "�",
        ),
        fc.constantFrom(
          "",
          "a",
          "č",
          "1",
          "_",
          "😀",
          "\u0301",
          ")",
          "s",
          "es",
          "\ud800",
          "\udc00",
          "�",
        ),
        (locale, needle, before, after) => {
          const target = `${before}${needle}${after}`;
          const prefix = locale === "ar" ? "[وفبكلس]*(?:ال)?" : "";
          const suffix = locale === "en" ? "(?:e?s)?" : "";
          const expected = new RegExp(
            `(?<!\\p{L})${prefix}${escape(needle.toWellFormed())}${suffix}(?!\\p{L})`,
            "iu",
          ).test(target.toWellFormed())
            ? [needle]
            : [];
          expect(findForbiddenTerms("Matter", target, locale, makeRules(needle))).toEqual(expected);
        },
      ),
      { seed: 27017, numRuns: 300 },
    );
  });
});
