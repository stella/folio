---
"@stll/folio-core": patch
---

A tracked batch that formats a container's last paragraph and also deletes every paragraph now refuses the deletion of that paragraph, as a direct batch does. It used to report both as applied and lose the format on accept.
