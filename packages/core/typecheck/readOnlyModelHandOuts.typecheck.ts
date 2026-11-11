import type { TableSelectionManager } from "../src/managers/TableSelectionManager";
import type { Document } from "../src/types/document";
import type { ReadonlyDocument } from "../src/types/documentModel";
import { resolveEffectiveSectionProperties } from "../src/utils/headerFooter";

// A reader that hands out a piece of the model hands it out read-only: the
// piece is the model's own, so a write through it would write the model.

export const selectedTable = (
  manager: TableSelectionManager,
  model: ReadonlyDocument,
): void => {
  const context = manager.selectCell(model, {
    tableIndex: 0,
    rowIndex: 0,
    columnIndex: 0,
  });
  const { table } = manager.getSnapshot();
  if (!context || !table) return;
  // @ts-expect-error the context's table is the model's own table
  context.table.rows.length = 0;
  // @ts-expect-error the selection's table is the model's own table
  table.rows.length = 0;
};

export const effectiveSectionProperties = (
  model: ReadonlyDocument,
  document: Document,
): void => {
  const properties = resolveEffectiveSectionProperties(
    model.package.document,
    false,
  );
  if (properties) {
    // @ts-expect-error a model's section properties are the model's own
    properties.marginTop = 0;
  }
  // A plain document still gives plain section properties.
  const plain = resolveEffectiveSectionProperties(
    document.package.document,
    false,
  );
  if (plain) plain.marginTop = 0;
};
