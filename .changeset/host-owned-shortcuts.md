---
"@stll/folio-core": minor
"@stll/folio-react": minor
"@stll/folio-vue": minor
---

`DocxEditor` takes `hostShortcuts` so a host with its own undo stack or print command owns those keys: `"history"` leaves Mod-z / Mod-y / Mod-Shift-z unbound, `"print"` leaves Cmd/Ctrl+P alone. `createStarterKit` takes the matching `historyShortcuts: "editor" | "host"`.
