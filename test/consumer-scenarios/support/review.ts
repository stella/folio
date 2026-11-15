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
  return settledText(
    await openReviewer(new Uint8Array(await reviewer.toBuffer())),
  );
};

/** Each tracked wrapper's kind and the words it holds, in order. */
const piecesOf = (node: unknown, wrapper: string): string[] => {
  if (typeof node !== "object" || node === null) return [];
  const record = node as Record<string, unknown>;
  const kind =
    record["type"] === "insertion" || record["type"] === "deletion"
      ? String(record["type"])
      : wrapper;
  if (record["type"] === "text" && typeof record["text"] === "string")
    return [`${kind}: ${record["text"]}`];
  const content = record["content"];
  return Array.isArray(content)
    ? content.flatMap((child) => piecesOf(child, kind))
    : [];
};

/**
 * A redline's body as written, one line per block: a paragraph's id, style,
 * its mark's tracked kind, its property changes and its words by tracked
 * kind, so a failed resolution shows the records it resolved.
 */
export const writtenLayout = async (bytes: Uint8Array): Promise<string> => {
  const reviewer = await openReviewer(bytes);
  return reviewer
    .toDocument()
    .package.document.content.map((block) => {
      if (block.type !== "paragraph") return block.type;
      const record = block as unknown as Record<string, unknown>;
      const mark = record["pPrMark"] as { kind?: unknown } | undefined;
      const changes = record["propertyChanges"];
      return JSON.stringify({
        paraId: block.paraId ?? null,
        styleId: block.formatting?.styleId ?? null,
        mark: mark?.kind ?? null,
        propertyChanges: Array.isArray(changes) ? changes.length : 0,
        pieces: piecesOf(block, "kept"),
      });
    })
    .join("\n");
};
