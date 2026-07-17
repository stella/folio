import { hasCjk } from "../../utils/scriptSegments";
import type { ParagraphAttrs, TextRun } from "../types";
import type { LineBreakPolicy } from "./lineBreakProvider";

const DEFAULT_HYPHENATION_ZONE_TWIPS = 360;

type EffectiveAutomaticHyphenation =
  | Readonly<{ type: "disabled" }>
  | Readonly<{
      type: "enabled";
      consecutiveLineLimit: number;
      hyphenationZoneTwips: number;
    }>;

type EffectiveLineBreakPolicy = Readonly<{
  /** Exact provider inputs; omitted OOXML values stay omitted for custom providers. */
  provider: Readonly<LineBreakPolicy>;
  /** `w:overflowPunct` defaults to enabled when omitted. */
  hangingPunctuation: boolean;
  /** Effective document and paragraph automatic-hyphenation state. */
  automaticHyphenation: EffectiveAutomaticHyphenation;
}>;

type ResolveEffectiveLineBreakPolicyOptions = {
  attrs: ParagraphAttrs | undefined;
  run: TextRun;
};

/** Resolve all authored line-breaking inputs used while measuring one text run. */
export const resolveEffectiveLineBreakPolicy = ({
  attrs,
  run,
}: ResolveEffectiveLineBreakPolicyOptions): EffectiveLineBreakPolicy => {
  const locale = resolveRunLocale(run);
  const rules = attrs?.lineBreakRules;
  const before = rules?.noLineBreaksBefore;
  const after = rules?.noLineBreaksAfter;
  const automaticHyphenation = attrs?.automaticHyphenation;

  const provider: LineBreakPolicy = {
    ...(locale ? { locale } : {}),
    ...(attrs?.kinsoku !== undefined ? { kinsoku: attrs.kinsoku } : {}),
    ...(before && languageMatches(before.language, locale)
      ? { noLineBreaksBefore: toCodePointSet(before.characters) }
      : {}),
    ...(after && languageMatches(after.language, locale)
      ? { noLineBreaksAfter: toCodePointSet(after.characters) }
      : {}),
    ...(rules?.useLegacyEthiopicAmharicRules ? { useLegacyEthiopicAmharicRules: true } : {}),
    ...(automaticHyphenation?.doNotHyphenateCaps !== undefined
      ? { doNotHyphenateCaps: automaticHyphenation.doNotHyphenateCaps }
      : {}),
    ...(run.allCaps === true ? { renderedAllCaps: true } : {}),
  };

  return {
    provider,
    hangingPunctuation: attrs?.overflowPunctuation !== false,
    automaticHyphenation: resolveAutomaticHyphenation(attrs),
  };
};

/** Cache fragments for every paragraph attribute consumed by the resolver. */
export const lineBreakPolicyCacheParts = (attrs: ParagraphAttrs): readonly string[] => {
  const parts: string[] = [];
  if (attrs.kinsoku !== undefined) {
    parts.push(`kinsoku:${attrs.kinsoku}`);
  }
  if (attrs.overflowPunctuation !== undefined) {
    parts.push(`overflow-punct:${attrs.overflowPunctuation}`);
  }
  if (attrs.suppressAutoHyphens !== undefined) {
    parts.push(`suppress-auto-hyphens:${attrs.suppressAutoHyphens}`);
  }

  const automaticHyphenation = attrs.automaticHyphenation;
  if (automaticHyphenation) {
    parts.push(
      `auto-hyphens:${automaticHyphenation.doNotHyphenateCaps}|${automaticHyphenation.consecutiveLineLimit}|${automaticHyphenation.hyphenationZoneTwips}`,
    );
  }

  const rules = attrs.lineBreakRules;
  if (rules) {
    parts.push(
      `line-break-rules:${rules.noLineBreaksBefore?.language}|${rules.noLineBreaksBefore?.characters}|${rules.noLineBreaksAfter?.language}|${rules.noLineBreaksAfter?.characters}|${rules.useLegacyEthiopicAmharicRules}`,
    );
  }
  return parts;
};

const resolveAutomaticHyphenation = (
  attrs: ParagraphAttrs | undefined,
): EffectiveAutomaticHyphenation => {
  const automaticHyphenation = attrs?.automaticHyphenation;
  if (automaticHyphenation?.enabled !== true || attrs?.suppressAutoHyphens === true) {
    return { type: "disabled" };
  }
  return {
    type: "enabled",
    consecutiveLineLimit: automaticHyphenation.consecutiveLineLimit ?? 0,
    hyphenationZoneTwips:
      automaticHyphenation.hyphenationZoneTwips ?? DEFAULT_HYPHENATION_ZONE_TWIPS,
  };
};

const resolveRunLocale = (run: TextRun): string | undefined => {
  const language = run.language;
  if (hasCjk(run.text)) {
    return language?.eastAsia ?? language?.val;
  }
  if (run.rtl) {
    return language?.bidi ?? language?.val;
  }
  return language?.val;
};

const languageMatches = (ruleLanguage: string | undefined, locale: string | undefined): boolean => {
  if (!ruleLanguage) {
    return true;
  }
  if (!locale) {
    return false;
  }
  const rule = ruleLanguage.toLowerCase();
  const active = locale.toLowerCase();
  return active === rule || active.startsWith(`${rule}-`) || rule.startsWith(`${active}-`);
};

/**
 * Split an already length-capped kinsoku character list (settingsParser.ts)
 * into a code-point Set so line-edge membership checks are O(1) instead of
 * an O(n) `String.includes` scan per character measured.
 */
const toCodePointSet = (characters: string): ReadonlySet<string> => new Set(Array.from(characters));
