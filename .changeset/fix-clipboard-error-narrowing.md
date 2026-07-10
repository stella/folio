---
"@stll/folio-core": patch
---

Narrow clipboard copy/read failures with an `instanceof Error` check instead of an unchecked `as Error` cast, so `onError` callbacks always receive a real `Error` even when a non-Error value is thrown.
