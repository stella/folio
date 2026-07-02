// Tailwind config for the self-sufficient `standalone.css` build only.
//
// The primary `editor.css` path does NOT use this: a host app runs its own
// Tailwind, scans folio's dist for utility classes, and supplies the semantic
// design tokens. `standalone.css` is the second path — it ships a pre-compiled
// copy of every utility folio's own components use so a consumer needs no
// Tailwind pipeline at all.
//
// `important: ".folio-root"` scopes every generated utility under the editor
// root (Tailwind v4 nests them as `.folio-root .<utility>`), so the compiled
// utilities can never leak into and restyle the host app. This is the v4
// equivalent of the v3 selector-string `important` option, loaded via the
// `@config` directive in `src/styles/standalone.css` (v4 dropped the CSS-side
// selector form; a legacy JS config still honors it). Pattern used by the
// upstream docx-editor's `.ep-root` scoping.
/** @type {import('tailwindcss').Config} */
export default {
  important: ".folio-root",
};
