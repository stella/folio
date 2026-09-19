/**
 * Whether `<w:t>` needs `xml:space="preserve"` to render the text as written.
 *
 * Without the attribute an XML reader trims leading and trailing spaces and
 * collapses runs of them, so the attribute is a pure function of the text: it
 * is needed exactly when that normalisation would change what the reader sees.
 * The model therefore keeps no flag of its own — a stored copy of a derived
 * fact only drifts, and it did: the editor round trip dropped it whenever two
 * runs merged, because ProseMirror has nowhere to carry it.
 */
export const requiresXmlSpacePreserve = (text: string): boolean =>
  text.startsWith(" ") || text.endsWith(" ") || text.includes("  ");
