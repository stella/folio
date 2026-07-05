/**
 * `@stll/folio-vue/composables`
 *
 * Vue composables mirroring the React `hooks` subpath — history, table
 * selection, find/replace, clipboard, zoom, tracked-changes, visual-line
 * navigation, and the high-level `useDocxEditor` host composable. Only the
 * composables the fork has ported are re-exported (React's `useAutoSave` and
 * `useCommentSidebarItems` have no Vue equivalent yet).
 *
 * @example
 * ```ts
 * import { useHistory, useFindReplace } from "@stll/folio-vue/composables";
 * ```
 *
 * @packageDocumentation
 * @public
 */

export * from "./useClipboard";
export * from "./useDocxEditor";
export * from "./useDragAutoScroll";
export * from "./useFindReplace";
export * from "./useFixedDropdown";
export * from "./useHistory";
export * from "./useSelectionHighlight";
export * from "./useTableResize";
export * from "./useTableSelection";
export * from "./useTrackedChanges";
export * from "./useVisualLineNavigation";
export * from "./useWheelZoom";
export * from "./useZoom";
