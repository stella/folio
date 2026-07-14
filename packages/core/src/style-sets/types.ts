import type {
  DocumentSettings,
  FontTable,
  NumberingDefinitions,
  SectionProperties,
  StyleDefinitions,
  Theme,
} from "../types/document";

export const DOCUMENT_STYLE_SET_VERSION = 1 as const;
export const DOCUMENT_PRESET_VERSION = 1 as const;

/**
 * A content-free, portable set of DOCX formatting resources.
 *
 * Source document content, properties, relationships, media, comments, and
 * revision data are deliberately excluded. Numbering and theme resources live
 * beside styles because paragraph styles can depend on them for their meaning.
 */
export type DocumentStyleSet = {
  version: typeof DOCUMENT_STYLE_SET_VERSION;
  name: string;
  initialParagraphStyleId: string;
  styles: StyleDefinitions;
  numbering?: NumberingDefinitions;
  theme?: Theme;
  fontTable?: FontTable;
  settings?: DocumentSettings;
};

/** Page-level choices are separate from reusable style resources. */
export type DocumentPreset = {
  version: typeof DOCUMENT_PRESET_VERSION;
  name: string;
  styleSet: DocumentStyleSet;
  sectionProperties: SectionProperties;
};
