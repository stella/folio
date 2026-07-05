// Barrel for the ProseMirror extension layer. Adapters import the manager and
// starter-kit factory from `@stll/folio-core/prosemirror/extensions`; individual
// node/mark extensions stay reachable at their explicit subpaths via the core
// `"./*"` export wildcard.
export { ExtensionManager } from "./ExtensionManager";
export { createStarterKit, type StarterKitOptions } from "./StarterKit";
