# Collaboration demo

Open the playground with `?collaboration=1`:

```text
http://localhost:4200/?collaboration=1
```

## What it shows

- **Yjs body sync** via folio's `collaboration` prop (`y-prosemirror` plugins)
- **Comment thread sync** via controlled `comments` + `onCommentsChange` backed by a `Y.Array`
- **WebRTC** peer discovery through public Yjs signaling servers (demo only)

## Try it

1. Start the playground: `bun --filter @stll/playground dev`
2. Open `/?collaboration=1` in two browser windows
3. Copy the share link (includes the `#room-…` hash) into the second window
4. Type in one window; edits and comments should appear in the other

Each tab gets a random display name persisted in `sessionStorage` for the session.
