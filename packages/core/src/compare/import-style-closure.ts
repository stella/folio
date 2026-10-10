import { getFolioDocxComparisonAccess, type FolioDocxReviewer } from "../ai-edits/headless";
import {
  remapFolioAIEditSnapshotNumberingReferences,
  remapFolioAIEditSnapshotStyleReferences,
} from "../ai-edits/snapshot";
import type { FolioAIEditSnapshot } from "../ai-edits/types";
import { createStyleResolver } from "../prosemirror/styles/styleResolver";

type ImportStyleClosureWithNumberingOptions = {
  destination: FolioDocxReviewer;
  source: FolioDocxReviewer;
  snapshots: readonly FolioAIEditSnapshot[];
  importedHeaderFooterSnapshots: readonly FolioAIEditSnapshot[];
  numberingReferences: readonly { numId: number; level: number }[];
};

/** Import numbering before styles bind to it, then rebind the source snapshots once. */
export const importStyleClosureWithNumbering = ({
  destination,
  source,
  snapshots,
  importedHeaderFooterSnapshots,
  numberingReferences,
}: ImportStyleClosureWithNumberingOptions) => {
  const access = getFolioDocxComparisonAccess(destination);
  const numbering = getFolioDocxComparisonAccess(source).numberingDefinitions();
  const numberingReferenceMap = access.planTargetNumberingReferences(
    numbering,
    numberingReferences,
  );
  const numberingStage: ReturnType<typeof access.stageTargetNumbering> =
    numberingReferenceMap === null
      ? "conflict"
      : access.stageTargetNumbering(numbering, numberingReferences, numberingReferenceMap);
  const styleImport = access.stageTargetStyles({
    source,
    snapshots,
    importedHeaderFooterSnapshots,
    ...(numberingReferenceMap !== null && { numberingReferenceMap }),
  });
  const reboundSnapshots = snapshots.map((snapshot) => {
    const styleRebound =
      styleImport.status === "unalignable"
        ? snapshot
        : remapFolioAIEditSnapshotStyleReferences({
            snapshot,
            styleIdMap: styleImport.styleIdMap,
            defaultParagraphStyleId: styleImport.defaultParagraphStyleId,
            importedStyleResolver: createStyleResolver(styleImport.styles),
            reconcileAuthoredFormatting: true,
          });
    return numberingReferenceMap === null
      ? styleRebound
      : remapFolioAIEditSnapshotNumberingReferences(styleRebound, numberingReferenceMap);
  });
  return { styleImport, numberingStage, snapshots: reboundSnapshots };
};
