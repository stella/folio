import type { Paragraph } from "../types/document";

/** Plain comment text has no authored paragraph formatting. */
export const commentTextParagraph = (text: string): Paragraph => ({
  type: "paragraph",
  content: [{ type: "run", formatting: {}, content: [{ type: "text", text }] }],
});
