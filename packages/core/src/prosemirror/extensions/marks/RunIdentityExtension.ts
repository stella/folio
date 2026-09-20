/**
 * Editor-side identity for one authored `w:r`, and the payload that element
 * carried which no formatting mark holds.
 *
 * Ordinary formatting marks cannot tell two adjacent source runs with
 * identical properties apart, so the save leg would fold them into one and
 * discard whichever record it did not keep. This mark is what keeps them
 * apart: it is in `getMarksKey`, so a change of identity is a run boundary,
 * and `joinsOwnedSourceRun` rejoins the leaves a page break cut out of one
 * authored run.
 */

import { expectRunIdentityMarkAttrs } from "../../attrs";
import { RUN_IDENTITY_ATTRIBUTE, RUN_IDENTITY_MARK_NAME, runIdentityAttrs } from "../../runIdentity";
import { createMarkExtension } from "../create";

const CANONICAL_RUN_IDENTITY_ID = /^(?:0|[1-9]\d*)$/u;

export const RunIdentityExtension = createMarkExtension({
  name: RUN_IDENTITY_MARK_NAME,
  schemaMarkName: RUN_IDENTITY_MARK_NAME,
  markSpec: {
    attrs: { id: {}, preservedAttributes: { default: null }, preserved: { default: null } },
    inclusive: false,
    parseDOM: [
      {
        tag: `span[${RUN_IDENTITY_ATTRIBUTE}]`,
        getAttrs(dom) {
          const rawId = dom.dataset["docxRunIdentity"];
          if (rawId === undefined || !CANONICAL_RUN_IDENTITY_ID.test(rawId)) {
            return false;
          }
          const id = Number(rawId);
          return Number.isSafeInteger(id) ? runIdentityAttrs(id) : false;
        },
      },
    ],
    // The id and nothing else. An rsid names a session listed in the source
    // package's `settings.xml`, and folio cannot merge rsid tables, so a span
    // pasted into another document must not claim the session it came from. It
    // keys differently from its source for that reason and becomes its own run
    // with no attributes, which is the answer typing already gets.
    toDOM(mark) {
      const { id } = expectRunIdentityMarkAttrs(mark);
      return ["span", { [RUN_IDENTITY_ATTRIBUTE]: String(id) }, 0];
    },
  },
});
