---
"@stll/folio-core": patch
---

Encode bytes to base64 through one owner. Four call sites each built `btoa`'s binary string their own way, and the markdown renderer's way was wrong: `TextDecoder("latin1")` is the windows-1252 decoder by specification, so any byte in 0x80-0x9F produced a character `btoa` rejects and registering an ordinary image threw `InvalidCharacterError` in browsers, where no `Buffer` fallback hides it. `utils/base64` now encodes bytes directly, using the runtime's `Uint8Array.prototype.toBase64` where there is one, and a lint rule keeps `btoa` out of package source.
