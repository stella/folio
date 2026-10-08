import type {
  DocumentBody,
  Endnote,
  Footnote,
  HeaderFooter,
  Table,
  TextBox,
} from "../types/document";
import { blockPlainText } from "./blockPlainText";

/** Read the document body with pending revisions accepted. */
export const getDocumentText = (body: DocumentBody): string => blockPlainText(body.content);

/** Approximate word count of the accepted body text. */
export const getWordCount = (body: DocumentBody): number => {
  const words = getDocumentText(body).trim().split(/\s+/u);
  return words.length > 0 && words[0] !== "" ? words.length : 0;
};

/** Character count of the accepted body text. */
export const getCharacterCount = (body: DocumentBody): number => getDocumentText(body).length;

/** Read an accepted table, including nested blocks and tracked structure. */
export const getTableText = (table: Table): string => blockPlainText([table]);

/** Read accepted text-box content for search and indexing. */
export const getTextBoxText = (textBox: TextBox): string => blockPlainText(textBox.content);

/** Read an accepted header or footer through the same walk as body and notes. */
export const getHeaderFooterText = (headerFooter: HeaderFooter): string =>
  blockPlainText(headerFooter.content);

/** Read accepted footnote content, including nested blocks. */
export const getFootnoteText = (footnote: Footnote): string => blockPlainText(footnote.content);

/** Read accepted endnote content, including nested blocks. */
export const getEndnoteText = (endnote: Endnote): string => blockPlainText(endnote.content);
