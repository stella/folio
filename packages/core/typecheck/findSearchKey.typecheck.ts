import type { FindOptions } from "../src/utils/findReplace";

// Bind the runtime option mutation matrix to every production matching option.
export const FIND_SEARCH_KEY_OPTION_CHANGES = {
  matchCase: { matchCase: true, matchWholeWord: false },
  matchWholeWord: { matchCase: false, matchWholeWord: true },
  useRegex: { matchCase: false, matchWholeWord: false, useRegex: true },
  displayMode: {
    matchCase: false,
    matchWholeWord: false,
    displayMode: "no-markup",
  },
  includeComments: {
    matchCase: false,
    matchWholeWord: false,
    includeComments: true,
  },
} satisfies Record<keyof FindOptions, FindOptions>;
