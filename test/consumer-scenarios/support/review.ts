/** Views of a reviewer's document for comparing review outcomes. */

import {
  FolioDocxReviewer,
  isFolioAIContentBlock,
} from "@stll/folio-core/server";

import { openReviewer, toArrayBuffer } from "./documents.ts";

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

/** One body block: a paragraph's every field but its content, and its words by tracked kind. */
const blockLine = (block: unknown): string => {
  const record = block as Record<string, unknown>;
  if (record["type"] !== "paragraph") return String(record["type"]);
  const { content: _content, ...fields } = record;
  return JSON.stringify({ ...fields, pieces: piecesOf(block, "kept") });
};

const bodyLines = (reviewer: Reviewer): string[] =>
  reviewer.toDocument().package.document.content.map(blockLine);

/** A version's body as the scenario's engine reads it, one line per block. */
export const bodyLayout = async (bytes: Uint8Array): Promise<string> =>
  bodyLines(await openReviewer(bytes)).join("\n");

/**
 * A redline as each engine reads it, before and after accepting every
 * change: one line per body block with its paragraph properties, its mark's
 * tracked record, its property changes, its section properties and its words
 * by tracked kind, then what a reader reads. A failed resolution shows the
 * records it resolved.
 */
export const writtenLayout = async (bytes: Uint8Array): Promise<string> => {
  const sections: string[] = [];
  for (const session of ["legacy", "canonical"] as const) {
    const reviewer = await FolioDocxReviewer.fromBuffer(toArrayBuffer(bytes), {
      session,
    });
    sections.push(
      `${session}, as written:`,
      ...bodyLines(reviewer),
      ...settledText(reviewer),
    );
    reviewer.acceptAll();
    sections.push(
      `${session}, accepted:`,
      ...bodyLines(reviewer),
      ...settledText(reviewer),
    );
  }
  return sections.join("\n");
};
