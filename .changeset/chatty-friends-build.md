---
---

Differential test harness gains the Open XML SDK reference parser. The only
`packages/core/src` change is a test file (`docx/__tests__/differential.test.ts`);
the projector and CI wiring live under `scripts/` and are dev/CI-only, so no
published-source behaviour changes and no release is needed.
