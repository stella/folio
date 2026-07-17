---
"@stll/folio-core": patch
---

Stop hidden and active DOCX content from leaking through save and read paths.
Selective save now bails to a full repack when the package contains
non-preservable entries (e.g. `word/vbaProject.bin`, embedded binaries) instead
of round-tripping them; hidden table-row text and `w:vanish` runs are excluded
from the AI snapshot; footnotes referenced only from hidden rows are no longer
painted; metadata-privacy scrubbing matches `docProps/core.xml` case-insensitively;
server text extraction resolves referenced headers/footers via relationships
instead of reading orphan parts; bound content-control clicks no longer throw;
and text-box anchor markers are stripped from pasted HTML and salted with a
per-load nonce.
