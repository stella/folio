// Core
import { DocExtension } from "../prosemirror/extensions/core/DocExtension";
import {
  historyKeyboardShortcuts,
  type HistoryShortcutOwner,
  redoCommand,
  undoCommand,
} from "../prosemirror/extensions/core/historyCommands";
import { ParagraphExtension } from "../prosemirror/extensions/core/ParagraphExtension";
import { TextExtension } from "../prosemirror/extensions/core/TextExtension";
import { createExtension } from "../prosemirror/extensions/create";
import { AutoBidiDetectionExtension } from "../prosemirror/extensions/features/AutoBidiDetectionExtension";
import { BaseKeymapExtension } from "../prosemirror/extensions/features/BaseKeymapExtension";
// oxlint-disable-next-line import/no-cycle -- BidiShortcutExtension reaches the singleton schema for runtime command lookup
import { BidiShortcutExtension } from "../prosemirror/extensions/features/BidiShortcutExtension";
import { ContentControlWidgetsExtension } from "../prosemirror/extensions/features/ContentControlWidgetsExtension";
import { DropCursorExtension } from "../prosemirror/extensions/features/DropCursorExtension";
import { EmptyParagraphFormatExtension } from "../prosemirror/extensions/features/EmptyParagraphFormatExtension";
import { GapCursorExtension } from "../prosemirror/extensions/features/GapCursorExtension";
import { ImageDragExtension } from "../prosemirror/extensions/features/ImageDragExtension";
import { ImagePasteExtension } from "../prosemirror/extensions/features/ImagePasteExtension";
import { JoinedRunStyleExtension } from "../prosemirror/extensions/features/JoinedRunStyleExtension";
// Features
import { ListExtension } from "../prosemirror/extensions/features/ListExtension";
import { PasteCleanupExtension } from "../prosemirror/extensions/features/PasteCleanupExtension";
import { PasteStyleInlinerExtension } from "../prosemirror/extensions/features/PasteStyleInlinerExtension";
import { SelectionTrackerExtension } from "../prosemirror/extensions/features/SelectionTrackerExtension";
// Marks
import {
  MARK_EXTENSIONS,
  MARK_NESTING_ORDER,
} from "../prosemirror/extensions/markRegistry";
import { BlockBookmarkBoundaryExtension } from "../prosemirror/extensions/nodes/BlockBookmarkBoundaryExtension";
import { BlockCommentRangeBoundaryExtension } from "../prosemirror/extensions/nodes/BlockCommentRangeBoundaryExtension";
import { BlockCustomXmlExtension } from "../prosemirror/extensions/nodes/BlockCustomXmlExtension";
import { BlockMoveRangeBoundaryExtension } from "../prosemirror/extensions/nodes/BlockMoveRangeBoundaryExtension";
import { BlockSdtExtension } from "../prosemirror/extensions/nodes/BlockSdtExtension";
import { BookmarkBoundaryExtension } from "../prosemirror/extensions/nodes/BookmarkBoundaryExtension";
import { CommentRangeBoundaryExtension } from "../prosemirror/extensions/nodes/CommentRangeBoundaryExtension";
import { CommentReferenceExtension } from "../prosemirror/extensions/nodes/CommentReferenceExtension";
import {
  FieldExtension,
  StructuredFieldExtension,
} from "../prosemirror/extensions/nodes/FieldExtension";
// Nodes
import { HardBreakExtension } from "../prosemirror/extensions/nodes/HardBreakExtension";
import { HorizontalRuleExtension } from "../prosemirror/extensions/nodes/HorizontalRuleExtension";
import { ImageExtension } from "../prosemirror/extensions/nodes/ImageExtension";
import { MathExtension } from "../prosemirror/extensions/nodes/MathExtension";
import { MoveRangeBoundaryExtension } from "../prosemirror/extensions/nodes/MoveRangeBoundaryExtension";
import { NoteMarkerExtension } from "../prosemirror/extensions/nodes/NoteMarkerExtension";
import { PageBreakExtension } from "../prosemirror/extensions/nodes/PageBreakExtension";
import { PageBreakRunExtension } from "../prosemirror/extensions/nodes/PageBreakRunExtension";
import { PreservedBlockExtension } from "../prosemirror/extensions/nodes/PreservedBlockExtension";
import { PreservedXmlExtension } from "../prosemirror/extensions/nodes/PreservedXmlExtension";
import { RangeAnchorExtension } from "../prosemirror/extensions/nodes/RangeAnchorExtension";
import { RenderedPageBreakExtension } from "../prosemirror/extensions/nodes/RenderedPageBreakExtension";
import { SdtExtension } from "../prosemirror/extensions/nodes/SdtExtension";
import { ShapeExtension } from "../prosemirror/extensions/nodes/ShapeExtension";
import { SymbolExtension } from "../prosemirror/extensions/nodes/SymbolExtension";
import { TabExtension } from "../prosemirror/extensions/nodes/TabExtension";
import { createTableExtensions } from "../prosemirror/extensions/nodes/TableExtension";
import { TextBoxAnchorExtension } from "../prosemirror/extensions/nodes/TextBoxAnchorExtension";
import { TextBoxExtension } from "../prosemirror/extensions/nodes/TextBoxExtension";
import type { AnyExtension } from "../prosemirror/extensions/types";
import type { SelectionChangeCallback } from "../prosemirror/plugins/selectionTracker";

/** Options for the DOCX editor extension set. */
export type StarterKitOptions = {
  /** Extensions to disable by name */
  disable?: string[];
  /** @deprecated The session owns history. This option has no effect. */
  historyDepth?: number;
  /** @deprecated The session owns history grouping. This option has no effect. */
  historyNewGroupDelay?: number;
  /** Who answers the undo and redo keys (default: `"editor"`). */
  historyShortcuts?: HistoryShortcutOwner;
  /** Selection change callback */
  onSelectionChange?: SelectionChangeCallback;
};

/**
 * Assemble the DOCX extension set shared with the reference editor.
 */
export function createStarterKitExtensions(
  options: StarterKitOptions = {},
): AnyExtension[] {
  const disabled = options.disable
    ? new Set(options.disable)
    : new Set<string>();

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
    createExtension({
      name: "history",
      onSchemaReady: () => ({
        commands: { undo: undoCommand, redo: redoCommand },
        ...(options.historyShortcuts !== "host" && {
          keyboardShortcuts: historyKeyboardShortcuts,
        }),
      }),
    })(),
  );

  // Marks. Registration order is DOM nesting order, outermost first, and
  // `MARK_NESTING_ORDER` is where that order is decided.
  for (const name of MARK_NESTING_ORDER) {
    add(name, MARK_EXTENSIONS[name]({ getInternalClipboardToken }));
  }

  // Nodes
  add(
    "bookmarkBoundary",
    BookmarkBoundaryExtension({ getInternalClipboardToken }),
  );
  add(
    "blockBookmarkBoundary",
    BlockBookmarkBoundaryExtension({ getInternalClipboardToken }),
  );
  add(
    "blockMoveRangeBoundary",
    BlockMoveRangeBoundaryExtension({ getInternalClipboardToken }),
  );
  add("blockCommentRangeBoundary", BlockCommentRangeBoundaryExtension());
  add(
    "commentReference",
    CommentReferenceExtension({ getInternalClipboardToken }),
  );
  add(
    "commentRangeBoundary",
    CommentRangeBoundaryExtension({ getInternalClipboardToken }),
  );
  add(
    "moveRangeBoundary",
    MoveRangeBoundaryExtension({ getInternalClipboardToken }),
  );
  add("rangeAnchor", RangeAnchorExtension({ getInternalClipboardToken }));
  add("hardBreak", HardBreakExtension());
  add("tab", TabExtension());
  add("symbol", SymbolExtension());
  add("preservedXml", PreservedXmlExtension());
  add("noteMarker", NoteMarkerExtension());
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
  add("field", FieldExtension({ getInternalClipboardToken }));
  add("field", StructuredFieldExtension({ getInternalClipboardToken }));
  add("sdt", SdtExtension());
  add("blockSdt", BlockSdtExtension());
  add("blockCustomXml", BlockCustomXmlExtension());
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
  add("joinedRunStyle", JoinedRunStyleExtension());
  add(
    "selectionTracker",
    options.onSelectionChange === undefined
      ? SelectionTrackerExtension()
      : SelectionTrackerExtension({
          onSelectionChange: options.onSelectionChange,
        }),
  );
  add("bidiShortcut", BidiShortcutExtension());
  add("autoBidiDetection", AutoBidiDetectionExtension());
  add("contentControlWidgets", ContentControlWidgetsExtension());

  return extensions;
}
