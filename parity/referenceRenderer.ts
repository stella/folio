/** External DOCX renderers that can serve as explicit comparison references. */

import {
  getLibreOfficePagePngs,
  getLibreOfficeTruth,
  getLibreOfficeVersion,
  isLibreOfficeAvailable,
} from "./libreOfficeTruth";
import type { DocGeom, ReferenceRendererId } from "./types";
import { getWordPagePngs, getWordTruth, getWordVersion, isWordAvailable } from "./wordTruth";

export type ReferenceRenderer = {
  id: ReferenceRendererId;
  displayName: string;
  installHint: string;
  isAvailable: () => Promise<boolean>;
  getVersion: () => Promise<string | null>;
  getTruth: (docxPath: string, options?: { refresh?: boolean }) => Promise<DocGeom>;
  getPagePngs: (docxPath: string, options?: { maxPages?: number }) => Promise<string[]>;
};

const LIBREOFFICE_RENDERER: ReferenceRenderer = {
  id: "libreoffice",
  displayName: "LibreOffice Writer",
  installHint: "Install LibreOffice from https://www.libreoffice.org/download/",
  isAvailable: isLibreOfficeAvailable,
  getVersion: getLibreOfficeVersion,
  getTruth: getLibreOfficeTruth,
  getPagePngs: getLibreOfficePagePngs,
};

const WORD_RENDERER: ReferenceRenderer = {
  id: "word",
  displayName: "Microsoft Word",
  installHint: "Install Word for Mac from https://www.microsoft.com/microsoft-365/word",
  isAvailable: isWordAvailable,
  getVersion: getWordVersion,
  getTruth: getWordTruth,
  getPagePngs: getWordPagePngs,
};

export const isReferenceRendererId = (value: string): value is ReferenceRendererId =>
  value === "libreoffice" || value === "word";

export const getReferenceRenderer = (id: ReferenceRendererId): ReferenceRenderer => {
  if (id === "libreoffice") return LIBREOFFICE_RENDERER;
  return WORD_RENDERER;
};
