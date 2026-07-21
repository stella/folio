import { toMarkdown } from "../../markdown";
import type { MarkdownOptions } from "../../markdown/types";
import { parseDocx } from "../parser";

/**
 * Convert a DOCX file to markdown in a single server-safe call.
 *
 * Composes {@link parseDocx} (bytes -> Document, with font preloading disabled
 * so it never touches the DOM) and {@link toMarkdown} (Document -> markdown).
 * This gives non-browser callers folio's full DOCX fidelity — headings,
 * tables, lists, and content controls (`w:sdt`) at every position — without
 * deep-importing internal subpaths or re-implementing an OOXML walker.
 */
export async function docxToMarkdown(
  input: ArrayBuffer | Uint8Array,
  opts?: MarkdownOptions,
): Promise<string> {
  const doc = await parseDocx(input, { preloadFonts: false });
  return toMarkdown(doc, opts);
}
