---
"@stll/folio-cli": minor
---

Add `folio render <file> -o <out.pdf|out.png|out.html> [--page n]`, which paints the document with folio's own layout and PDF or DOM backend (PNG through an optional `playwright-core` Chromium), and `folio serve <file> [--port n]`, a read-only live preview on 127.0.0.1 behind a random token that re-renders when the file's version changes.
