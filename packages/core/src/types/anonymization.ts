import type { TextPosition } from "@stll/docx-core/ops";

/** A term the host wants highlighted, with the surface forms it takes. */
export type FolioAnonymizationTerm = {
  /** The term's canonical surface form, reported back with each match. */
  canonical: string;
  /** Label slug (for example "person" or "organization"). */
  label: string;
  /** Other surface forms matched verbatim besides `canonical`. */
  variants?: readonly string[];
};

/** One occurrence of a term, as public text positions. */
export type FolioAnonymizationMatch = {
  start: TextPosition;
  end: TextPosition;
  label: string;
  canonical: string;
};
