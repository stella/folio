import { sectionReferenceHistory } from "../docx/sectionReferenceHistory";
import type { SectionProperties } from "../types/document";
import { canonicalJson } from "../utils/canonicalJson";
import type { FolioRevisionStamp } from "./applyTypes";

type StageFinalSectionPropertiesOptions = {
  target: SectionProperties;
  previous: SectionProperties;
  revision: FolioRevisionStamp;
  author: string;
};

/** The tracked final-section state both reviewer engines stage. */
export const stagedFinalSectionProperties = ({
  target,
  previous,
  revision,
  author,
}: StageFinalSectionPropertiesOptions): SectionProperties | null => {
  if (canonicalJson(target) === canonicalJson(previous)) return null;
  const previousReferences = sectionReferenceHistory({ previous, target });
  return {
    ...target,
    propertyChanges: [
      {
        type: "sectionPropertyChange",
        info: { id: revision.idSeed, author, date: revision.date },
        previousProperties: previous,
        ...(previousReferences !== undefined && { previousReferences }),
        currentProperties: target,
      },
    ],
  };
};
