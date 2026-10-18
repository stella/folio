import { AhoCorasick } from "@stll/aho-corasick";

/** Own matchers for one document or query run; discard the scope with that run. */
export const createLiteralSearchScope = () => {
  const matcherCache = new Map<string, AhoCorasick>();
  const literalMatcher = (patterns: string[]) => {
    const wellFormedPatterns = patterns.map((pattern) => pattern.toWellFormed());
    const key = JSON.stringify(wellFormedPatterns);
    let matcher = matcherCache.get(key);
    if (!matcher) {
      matcher = new AhoCorasick(wellFormedPatterns);
      matcherCache.set(key, matcher);
    }
    return matcher;
  };

  /** Compile once for the text-query configuration, then reuse across lines. */
  const createLiteralSearch = (needle: string) => {
    if (needle.length === 0) return (_text: string) => true;
    const matcher = literalMatcher([needle]);
    return (text: string) => matcher.isMatch(text.toWellFormed());
  };

  /** Both containment directions give an inspector candidate full similarity. */
  const createCandidateContainment = (query: string, candidates: readonly string[]) => {
    const matcher = literalMatcher(candidates.filter((candidate) => candidate.length > 0));
    const contained = new Set(
      matcher.findOverlappingIter(query.toWellFormed()).map(({ text }) => text),
    );
    contained.add("");
    const matchesQuery = createLiteralSearch(query);
    return (text: string) => matchesQuery(text) || contained.has(text.toWellFormed());
  };

  return { createLiteralSearch, createCandidateContainment };
};
