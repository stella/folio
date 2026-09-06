import { describe, expect, test } from "bun:test";

import { pairFolioDocumentStories } from "./document-stories";

const header = (relationshipId: string) => ({ type: "header", relationshipId }) as const;
const footer = (relationshipId: string) => ({ type: "footer", relationshipId }) as const;

describe("pairFolioDocumentStories", () => {
  test("pairs a header with the one carrying the same relationship id", () => {
    // Two revisions of one file carry their part ids forward, and matching on
    // them pairs the right header however the parts are ordered.
    expect(
      pairFolioDocumentStories(
        [{ type: "main" }, header("rId6"), header("rId7")],
        [{ type: "main" }, header("rId7"), header("rId6")],
      ),
    ).toEqual([
      { baseStory: { type: "main" }, revisedStory: { type: "main" } },
      { baseStory: header("rId6"), revisedStory: header("rId6") },
      { baseStory: header("rId7"), revisedStory: header("rId7") },
    ]);
  });

  test("pairs headers by order when the two packages share no relationship id", () => {
    // A relationship id names a part inside one package. Two independently
    // authored documents never agree on one, and pairing on it alone left
    // every header of such a comparison reported as present on one side only.
    expect(
      pairFolioDocumentStories([header("rId3"), footer("rId4")], [header("rId8"), footer("rId9")]),
    ).toEqual([
      { baseStory: header("rId3"), revisedStory: header("rId8") },
      { baseStory: footer("rId4"), revisedStory: footer("rId9") },
    ]);
  });

  test("a header pairs with a header, never with a footer", () => {
    expect(pairFolioDocumentStories([header("rId3")], [footer("rId9")])).toEqual([
      { baseStory: header("rId3"), revisedStory: null },
      { baseStory: null, revisedStory: footer("rId9") },
    ]);
  });

  test("the surplus stays unpaired when one side has more", () => {
    expect(pairFolioDocumentStories([header("rId3")], [header("rId8"), header("rId9")])).toEqual([
      { baseStory: header("rId3"), revisedStory: header("rId8") },
      { baseStory: null, revisedStory: header("rId9") },
    ]);
  });

  test("an identity match is never stolen by an earlier order match", () => {
    // The base's second header names the target's first by id. Pairing the
    // first base header against it by order would leave the exact match
    // unpaired and put two unrelated headers opposite each other.
    expect(
      pairFolioDocumentStories([header("rId3"), header("rId8")], [header("rId8"), header("rId9")]),
    ).toEqual([
      { baseStory: header("rId3"), revisedStory: header("rId9") },
      { baseStory: header("rId8"), revisedStory: header("rId8") },
    ]);
  });

  test("notes pair by note id and never by order", () => {
    // A note id is the note itself, not a package-scoped pointer: two
    // documents that both number a footnote 2 mean the same footnote.
    expect(
      pairFolioDocumentStories(
        [{ type: "footnote", noteId: 2 }],
        [{ type: "footnote", noteId: 3 }],
      ),
    ).toEqual([
      { baseStory: { type: "footnote", noteId: 2 }, revisedStory: null },
      { baseStory: null, revisedStory: { type: "footnote", noteId: 3 } },
    ]);
  });
});
