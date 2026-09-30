import { expect, test } from "bun:test";
import fc from "fast-check";

import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";
import { createDocx } from "../docx/rezip";
import { paragraph } from "../docx/server/build";
import { FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION } from "../document-operations";
import { createEmptyDocument } from "../utils/createDocument";
import { FolioDocxReviewer } from "./headless";

const revisionView = (reviewer: FolioDocxReviewer) =>
  reviewer
    .getChanges()
    .map(({ type, author, text, blockId }) => ({ type, author, text, blockId }))
    .toSorted((left, right) => JSON.stringify(left).localeCompare(JSON.stringify(right)));

// Earlier round-trip tests checked visibility and one nested deletion, but
// did not compare the full revision census after later paragraph deletion.
test(
  "tracked replacements followed by paragraph deletion preserve the revision census across save",
  async () => {
    await assertProperty(
      fc.asyncProperty(
        fc.array(fc.stringMatching(/^[A-Za-z]{1,12}$/u), { minLength: 1, maxLength: 4 }),
        fc.boolean(),
        fc.boolean(),
        async (words, reopenBetweenEdits, anotherAuthor) => {
          const document = createEmptyDocument();
          document.package.document.content = [
            { ...paragraph("Start clause ends."), paraId: "10000100" },
            { ...paragraph("Anchor."), paraId: "10000200" },
          ];
          let reviewer = await FolioDocxReviewer.fromBuffer(await createDocx(document), {
            author: "Reviewer",
          });
          let find = "clause";
          for (const [index, word] of words.entries()) {
            const block = reviewer.getContent().at(0);
            if (!block) throw new Error("the fixture paragraph is missing");
            const replacement = `${word}${index}`;
            const result = reviewer.applyDocumentOperations({
              version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
              mode: "tracked-changes",
              operations: [
                {
                  id: "replace",
                  type: "replaceInBlock",
                  blockId: block.id,
                  find,
                  replace: replacement,
                },
              ],
            });
            expect(result.applied).toHaveLength(1);
            const reopened = await FolioDocxReviewer.fromBuffer(await reviewer.toBuffer(), {
              author: anotherAuthor ? "Second Reviewer" : "Reviewer",
            });
            expect(revisionView(reopened)).toEqual(revisionView(reviewer));
            if (reopenBetweenEdits) reviewer = reopened;
            find = replacement;
          }
          const block = reviewer.getContent().at(0);
          if (!block) throw new Error("the fixture paragraph is missing");
          const result = reviewer.applyDocumentOperations({
            version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
            mode: "tracked-changes",
            operations: [{ id: "delete", type: "deleteBlock", blockId: block.id }],
          });
          expect(result.applied).toHaveLength(1);
          const reopened = await FolioDocxReviewer.fromBuffer(await reviewer.toBuffer());
          expect(revisionView(reopened)).toEqual(revisionView(reviewer));
        },
      ),
      { numRuns: 40 },
    );
  },
  propertyTestTimeout(20_000),
);
