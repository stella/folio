---
"@stll/folio-core": patch
---

Two page-break shapes Word writes routinely now convert. A page break opening a table cell projects as the row's break whatever else the cell holds, where before the cell had to contain that one paragraph alone. A paragraph carrying both a text-box anchor and a page break converts when the anchor precedes the break, which is the arrangement layout already projects faithfully; only an anchor that follows the break is still refused.
