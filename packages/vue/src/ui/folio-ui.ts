/**
 * Vue UI-injection layer (placeholder).
 *
 * folio lets a host override the editor's built-in chrome primitives (Button,
 * …) with its own design-system components. The React package models this as a
 * `FolioUIComponents` record of injectable components.
 *
 * TODO(vue, Phase E): replace this placeholder with the real injectable-component
 * contract (the concrete component map keyed by primitive name, typed against the
 * Vue component surface). Until then `DocxEditorProps.components` accepts an
 * open record so the props contract can mirror React's shape.
 */

// eslint-disable-next-line no-restricted-syntax -- placeholder until Phase E lands the real contract
export type FolioUIComponents = Record<string, unknown>;
