/** Exact section restoration can retain an explicitly undefined own field. */
import type { DocumentBody, Paragraph, Section } from "../../src/model/content";

const PARAGRAPH = {
  type: "paragraph",
  content: [],
  sectionProperties: undefined,
} satisfies Paragraph;

const SECTION = {
  properties: {},
  content: [],
  headers: undefined,
  footers: undefined,
} satisfies Section;

const BODY = {
  content: [],
  sections: undefined,
  finalSectionProperties: undefined,
} satisfies DocumentBody;

export type SectionPresenceProof = [typeof PARAGRAPH, typeof SECTION, typeof BODY];
