---
---

CI installs Bun from an Actions cache keyed by runner and the pinned version. A miss downloads it from the release host, or else from the npm registry, each checked against its published digest.
