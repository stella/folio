/**
 * The canonical save writes over the bytes its document was read from. Handed
 * them as the producer wrote them, with the main document at another path, it
 * saves one main document part.
 */

import { describe, test } from "bun:test";

import { createCanonicalSession } from "../canonical/canonicalSession";
import {
  expectOneMainDocument,
  openedAsWritten,
} from "../docx/__tests__/mainDocumentAtAnotherPath";
import { serializeCanonicalSave } from "../docx/canonicalSave";

describe("a canonical save handed the package as its producer wrote it", () => {
  test("the canonical save writes one main document", async () => {
    const { document } = await openedAsWritten();
    const session = createCanonicalSession(document).unwrap();
    const saved = await serializeCanonicalSave({
      snapshot: session.captureSaveSnapshot(),
      featureFlags: { selectiveSave: true },
    });
    await expectOneMainDocument(saved.buffer, "Authored text");
  });
});
