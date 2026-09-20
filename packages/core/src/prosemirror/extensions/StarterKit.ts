/**
 * StarterKit — bundles all extensions into a ready-to-use set
 *
 * Usage:
 *   const extensions = createStarterKit();
 *   const manager = new ExtensionManager(extensions);
 *   manager.buildSchema();
 *   manager.initializeRuntime();
 */

import type { SelectionChangeCallback } from "../plugins/selectionTracker";
// Core
import { DocExtension } from "./core/DocExtension";
import { HistoryExtension } from "./core/HistoryExtension";
import { ParagraphExtension } from "./core/ParagraphExtension";
import { TextExtension } from "./core/TextExtension";
import { AutoBidiDetectionExtension } from "./features/AutoBidiDetectionExtension";
import { BaseKeymapExtension } from "./features/BaseKeymapExtension";
// oxlint-disable-next-line import/no-cycle -- BidiShortcutExtension reaches the singleton schema for runtime command lookup
import { BidiShortcutExtension } from "./features/BidiShortcutExtension";
import { ContentControlWidgetsExtension } from "./features/ContentControlWidgetsExtension";
import { DropCursorExtension } from "./features/DropCursorExtension";
import { EmptyParagraphFormatExtension } from "./features/EmptyParagraphFormatExtension";
import { GapCursorExtension } from "./features/GapCursorExtension";
import { ImageDragExtension } from "./features/ImageDragExtension";
import { ImagePasteExtension } from "./features/ImagePasteExtension";
// Features
import { ListExtension } from "./features/ListExtension";
import { ParagraphChangeTrackerExtension } from "./features/ParagraphChangeTrackerExtension";
import { ParaIdAllocatorExtension } from "./features/ParaIdAllocatorExtension";
import { PasteCleanupExtension } from "./features/PasteCleanupExtension";
import { PasteStyleInlinerExtension } from "./features/PasteStyleInlinerExtension";
import { SelectionTrackerExtension } from "./features/SelectionTrackerExtension";
// Marks
import { MARK_EXTENSIONS } from "./markRegistry";
import { BlockSdtExtension } from "./nodes/BlockSdtExtension";
import { BookmarkBoundaryExtension } from "./nodes/BookmarkBoundaryExtension";
import { CommentReferenceExtension } from "./nodes/CommentReferenceExtension";
import { FieldExtension, StructuredFieldExtension } from "./nodes/FieldExtension";
// Nodes
import { HardBreakExtension } from "./nodes/HardBreakExtension";
import { HorizontalRuleExtension } from "./nodes/HorizontalRuleExtension";
import { ImageExtension } from "./nodes/ImageExtension";
import { MathExtension } from "./nodes/MathExtension";
import { PageBreakExtension } from "./nodes/PageBreakExtension";
import { PageBreakRunExtension } from "./nodes/PageBreakRunExtension";
import { RenderedPageBreakExtension } from "./nodes/RenderedPageBreakExtension";
import { SdtExtension } from "./nodes/SdtExtension";
import { ShapeExtension } from "./nodes/ShapeExtension";
import { PreservedBlockExtension } from "./nodes/PreservedBlockExtension";
import { PreservedXmlExtension } from "./nodes/PreservedXmlExtension";
import { SymbolExtension } from "./nodes/SymbolExtension";
import { TabExtension } from "./nodes/TabExtension";
import { createTableExtensions } from "./nodes/TableExtension";
import { TextBoxExtension } from "./nodes/TextBoxExtension";
import { TextBoxAnchorExtension } from "./nodes/TextBoxAnchorExtension";
import type { AnyExtension } from "./types";

export type StarterKitOptions = {
  /** Extensions to disable by name */
  disable?: string[];
  /** History depth (default: 100) */
  historyDepth?: number;
  /** History new group delay (default: 500) */
  historyNewGroupDelay?: number;
  /** Selection change callback */
  onSelectionChange?: SelectionChangeCallback;
};

/**
 * Create the full set of extensions for the DOCX editor
 */
export function createStarterKit(options: StarterKitOptions = {}): AnyExtension[] {
  const disabled = options.disable ? new Set(options.disable) : new Set<string>();

  const extensions: AnyExtension[] = [];
  let internalClipboardToken: string | undefined;
  const getInternalClipboardToken = (): string => {
    internalClipboardToken ??= globalThis.crypto.randomUUID();
    return internalClipboardToken;
  };

  function add(name: string, ext: AnyExtension): void {
    if (!disabled.has(name)) {
      extensions.push(ext);
    }
  }

  // Core (always included unless explicitly disabled)
  add("doc", DocExtension());
  add("text", TextExtension());
  add("paragraph", ParagraphExtension());
  add(
    "history",
    HistoryExtension({
      ...(options.historyDepth !== undefined ? { depth: options.historyDepth } : {}),
      ...(options.historyNewGroupDelay !== undefined
        ? { newGroupDelay: options.historyNewGroupDelay }
        : {}),
    }),
  );

  // Marks. The registry is the schema's mark set and its DOM nesting order;
  // the object's key order is the registration order.
  for (const [name, extension] of Object.entries(MARK_EXTENSIONS)) {
    add(name, extension());
  }

  // Nodes
  add("bookmarkBoundary", BookmarkBoundaryExtension({ getInternalClipboardToken }));
  add("commentReference", CommentReferenceExtension({ getInternalClipboardToken }));
  add("hardBreak", HardBreakExtension());
  add("tab", TabExtension());
  add("symbol", SymbolExtension());
  add("preservedXml", PreservedXmlExtension());
  add("image", ImageExtension());
  add("textBox", TextBoxExtension());
  add("textBoxAnchor", TextBoxAnchorExtension({ getInternalClipboardToken }));
  add("shape", ShapeExtension());
  add("imageDrag", ImageDragExtension());
  add("imagePaste", ImagePasteExtension());
  add("dropCursor", DropCursorExtension());
  add("gapCursor", GapCursorExtension());
  add("horizontalRule", HorizontalRuleExtension());
  add("pageBreak", PageBreakExtension());
  add("pageBreakRun", PageBreakRunExtension());
  add("renderedPageBreak", RenderedPageBreakExtension());
  add("field", FieldExtension());
  add("field", StructuredFieldExtension({ getInternalClipboardToken }));
  add("sdt", SdtExtension());
  add("blockSdt", BlockSdtExtension());
  add("preservedBlock", PreservedBlockExtension());
  add("math", MathExtension());

  // Table (5 extensions grouped)
  if (!disabled.has("table")) {
    extensions.push(...createTableExtensions());
  }

  // Features
  add("pasteCleanup", PasteCleanupExtension({ getInternalClipboardToken }));
  add("pasteStyleInliner", PasteStyleInlinerExtension());
  add("list", ListExtension());
  add("baseKeymap", BaseKeymapExtension());
  add("emptyParagraphFormat", EmptyParagraphFormatExtension());
  add(
    "selectionTracker",
    options.onSelectionChange === undefined
      ? SelectionTrackerExtension()
      : SelectionTrackerExtension({
          onSelectionChange: options.onSelectionChange,
        }),
  );
  // Register the paraId allocator BEFORE the change tracker so any
  // freshly-allocated id is already on the paragraph when the tracker
  // records it as changed. The plugin sets `addToHistory: false` so
  // undo/redo doesn't trip on the allocation.
  add("paraIdAllocator", ParaIdAllocatorExtension());
  add("paragraphChangeTracker", ParagraphChangeTrackerExtension());
  add("bidiShortcut", BidiShortcutExtension());
  add("autoBidiDetection", AutoBidiDetectionExtension());
  add("contentControlWidgets", ContentControlWidgetsExtension());

  return extensions;
}
