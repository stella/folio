import { panic } from "better-result";

import {
  lockedControlMessage,
  type LockedControlViolation,
} from "@stll/docx-core/model";

type ReviewLockFact = {
  id: number;
  suggestionId: string | null;
  violation: LockedControlViolation | undefined;
};

type ClassifyReviewLocksOptions<Source, Change extends ReviewLockFact> = {
  stories: readonly { source: Source; changes: readonly Change[] }[];
  grouping: "whole-proposal" | "individual-change";
};

/** Source-specific probes supply facts; a held proposal stays whole across all stories. */
export const classifyReviewLocks = <Source, Change extends ReviewLockFact>({
  stories,
  grouping,
}: ClassifyReviewLocksOptions<Source, Change>) => {
  const heldSuggestions = new Map<string, LockedControlViolation>();
  switch (grouping) {
    case "whole-proposal":
      for (const { changes } of stories)
        for (const { suggestionId, violation } of changes)
          if (suggestionId !== null && violation !== undefined)
            heldSuggestions.set(suggestionId, violation);
      break;
    case "individual-change":
      break;
    default: {
      const unhandled: never = grouping;
      return panic("Unknown review lock grouping", { grouping: unhandled });
    }
  }
  return {
    heldSuggestionIds: new Set(heldSuggestions.keys()),
    stories: stories.map(({ source, changes }) => ({
      source,
      changes: changes.map((change) => {
        const violation =
          change.violation ??
          (change.suggestionId === null
            ? undefined
            : heldSuggestions.get(change.suggestionId));
        if (violation === undefined) return { type: "free", change } as const;
        return {
          type: "held",
          change,
          violation,
          message:
            change.violation === undefined
              ? `Another change in this proposal is held: ${lockedControlMessage(violation)}`
              : lockedControlMessage(violation),
        } as const;
      }),
    })),
  };
};
