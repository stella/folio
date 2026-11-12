// A field-totality gate in a test file is read by no compiler, so the case
// per matching option lives here, where the package typecheck compiles it.
import type { FindOptions } from "@stll/folio-core/utils/findReplace";

import type { UseFindReplaceReturn } from "../useFindReplace";

type MatchingOptionFixture = {
  text: string;
  query: string;
  initialOptions: Partial<FindOptions>;
  change: (search: UseFindReplaceReturn) => void;
};

export const matchingOptionFixtures = {
  matchCase: {
    text: "Hit",
    query: "hit",
    initialOptions: {},
    change: (search) => {
      search.options.matchCase = true;
    },
  },
  matchWholeWord: {
    text: "hitched",
    query: "hit",
    initialOptions: {},
    change: (search) => {
      search.options.matchWholeWord = true;
    },
  },
  useRegex: {
    text: "hit",
    query: "^hit",
    initialOptions: { useRegex: true },
    change: (search) => {
      search.options.useRegex = false;
    },
  },
} satisfies Record<
  keyof Pick<FindOptions, "matchCase" | "matchWholeWord" | "useRegex">,
  MatchingOptionFixture
>;
