/** Views of a reviewer's document for comparing review outcomes. */

import { isFolioAIContentBlock } from "@stll/folio-core/server";

import { openReviewer } from "./documents.ts";

type Reviewer = Awaited<ReturnType<typeof openReviewer>>;

/** The blocks a reader is left with: kind and text, deleted blocks gone. */
export const settledText = (reviewer: Reviewer): string[] =>
  reviewer
    .getContent()
    .filter((block) => isFolioAIContentBlock(block))
    .filter((block) => block.text.length > 0)
    .map((block) => `${block.kind}: ${block.text}`);

/** Open `bytes`, resolve every change one way, and read what is left. */
export const resolvedText = async (
  bytes: Uint8Array,
  resolution: "accept" | "reject",
): Promise<string[]> => {
  const reviewer = await openReviewer(bytes);
  if (resolution === "accept") {
    reviewer.acceptAll();
  } else {
    reviewer.rejectAll();
  }
  // Resolve, save and reopen, so the answer is what the package says.
  return settledText(await openReviewer(new Uint8Array(await reviewer.toBuffer())));
};
