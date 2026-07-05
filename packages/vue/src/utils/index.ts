// View-layer UI utilities for the Vue adapter. These are DOM/UI-shaped helpers
// that stay in the adapter so `@stll/folio-core` remains framework-neutral.
// `findVerticalScrollParent` is re-exported from core (core's own
// visual-line-navigation depends on it, so it is the single source of truth).
export * from "./autoScroll";
export * from "./cardStyles";
export * from "./colorMode";
export * from "./comments";
export * from "./cssTypes";
export * from "./fontOptions";
export * from "./listState";
export * from "./reportIssue";
export * from "./selectionHighlight";
export * from "./sidebarConstants";
export * from "./stylePreview";
export * from "./textSelection";
export {
  findVerticalScrollParent,
  findVerticalScrollParentOrRoot,
} from "@stll/folio-core/utils/findVerticalScrollParent";
// TODO(vue): fontLoader (getRenderableDocumentFonts / loadFontDefinitions /
// FontDefinition …) is reconciled against core's embedded-font surface when the
// useFontLifecycle composable lands, to avoid duplicating core's font handling.
