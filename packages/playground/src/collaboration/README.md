# Collaboration demo

Open the playground with `?collaboration=1`:

```text
http://localhost:4200/?collaboration=1
```

## What it shows

- **A canonical room** through folio's `collaboration` prop: every edit is an
  operation of the room's journal, sequenced by one authority
- **The authority in one tab**: the tabs of this origin elect a leader with a
  Web Lock and talk over a `BroadcastChannel`; when the leader closes, another
  tab takes over from the journal
- **A durable journal** in IndexedDB: a reloaded tab, or a tab opened later,
  catches up from it
- **Comments in the journal**: threads, replies and resolution travel as
  operations with the text, so there is one comment list
- **Presence**: each tab's selection, drawn in the other tabs

## Try it

1. Start the playground: `bun --filter @stll/playground dev`
2. Open `/?collaboration=1` in a tab, then open the same URL (with its
   `#room-…` hash) in a second tab of the same browser
3. Type, suggest or comment in one tab; the other follows

`?file=<fixture>` opens another fixture (`sample.docx` by default). Each tab
gets a random display name persisted in `sessionStorage` for the session.
