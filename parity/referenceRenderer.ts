/** External DOCX renderers that can serve as explicit comparison references. */

import {
  getLibreOfficePagePngs,
  getLibreOfficeGeometry,
  getLibreOfficeVersion,
  isLibreOfficeAvailable,
} from "./libreOfficeReference";
import type { DocGeom, ReferenceRendererId, ReviewView } from "./types";
import { getWordPagePngs, getWordTruth, getWordVersion, isWordAvailable } from "./wordTruth";

export type ReferenceRenderer = {
  id: ReferenceRendererId;
  displayName: string;
  installHint: string;
  reviewViews: readonly ReviewView[];
  isAvailable: () => Promise<boolean>;
  getVersion: () => Promise<string | null>;
  getGeometry: (
    docxPath: string,
    options: { refresh?: boolean; reviewView: ReviewView },
  ) => Promise<DocGeom>;
  getPagePngs: (
    docxPath: string,
    options: { maxPages?: number; reviewView: ReviewView },
  ) => Promise<string[]>;
};

const LIBREOFFICE_RENDERER: ReferenceRenderer = {
  id: "libreoffice",
  displayName: "LibreOffice Writer",
  installHint: "Install LibreOffice from https://www.libreoffice.org/download/",
  reviewViews: ["default"],
  isAvailable: isLibreOfficeAvailable,
  getVersion: getLibreOfficeVersion,
  getGeometry: getLibreOfficeGeometry,
  getPagePngs: getLibreOfficePagePngs,
};

const WORD_RENDERER: ReferenceRenderer = {
  id: "word",
  displayName: "Microsoft Word",
  installHint: "Install Word for Mac from https://www.microsoft.com/microsoft-365/word",
  reviewViews: ["final", "all-markup"],
  isAvailable: isWordAvailable,
  getVersion: getWordVersion,
  getGeometry: getWordTruth,
  getPagePngs: getWordPagePngs,
};

export const isReferenceRendererId = (value: string): value is ReferenceRendererId =>
  value === "libreoffice" || value === "word";

export const getReferenceRenderer = (id: ReferenceRendererId): ReferenceRenderer => {
  switch (id) {
    case "libreoffice":
      return LIBREOFFICE_RENDERER;
    case "word":
      return WORD_RENDERER;
  }
};
