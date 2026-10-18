import { AhoCorasick } from "@stll/aho-corasick";

/** Compile a stable vocabulary; results retain its order and duplicates. */
export const createLiteralChecklist = (patterns: readonly string[]) => {
  const vocabulary = [...patterns];
  const nonempty = vocabulary.filter((pattern) => pattern.length > 0);
  const matcher = new AhoCorasick(nonempty);
  return (text: string) => {
    const found = new Set(
      matcher.findOverlappingIter(text).map(({ text: matchedText }) => matchedText),
    );
    found.add("");
    return {
      present: vocabulary.filter((pattern) => found.has(pattern)),
      missing: vocabulary.filter((pattern) => !found.has(pattern)),
    };
  };
};
