---
"@stll/folio-core": patch
---

Compare inline-atom topology on document facts alone. A text box projects a `textBoxAnchor` whose id is minted per conversion and salted with a random nonce, and the key both sides were compared on carried that id verbatim, so any paragraph holding a text box beside a field, image or page break failed to align — a document differed from itself. The key now drops attributes a conversion mints for itself, and `offsetAt` no longer claims a refusal it could not return.
