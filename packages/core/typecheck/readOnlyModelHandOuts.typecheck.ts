import type { FolioEditor } from "../src/controller/folioEditor";
import type { HeaderFooterEditorManager } from "../src/controller/headerFooterEditorManager";
import type { HiddenEditorApi } from "../src/controller/hiddenEditorApi";
import type { NoteEditorManager } from "../src/controller/noteEditorManager";
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

export const storySnapshots = (
  notes: NoteEditorManager,
  parts: HeaderFooterEditorManager,
  model: ReadonlyDocument,
): void => {
  const noted = notes.snapshotDocument(model);
  // @ts-expect-error a note snapshot shares what no edit touched with the model
  noted.package.document.content.length = 0;
  const parted = parts.snapshotDocument(model);
  // @ts-expect-error a part snapshot shares what no edit touched with the model
  parted.package.headers?.clear();
};

export const hiddenEditorDocument = (
  api: HiddenEditorApi,
  editor: FolioEditor,
): void => {
  const built = api.getDocument();
  if (built) {
    // @ts-expect-error the built document holds the model's own parts
    built.package.document.finalSectionProperties = undefined;
    // @ts-expect-error the built document holds the model's own styles
    built.package.styles?.styles.splice(0);
  }
  const controlled = editor.getDocument();
  if (controlled) {
    // @ts-expect-error the controller's document is the hidden editor's
    controlled.package.document.content.length = 0;
  }
};
