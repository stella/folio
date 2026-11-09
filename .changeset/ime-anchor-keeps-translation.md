---
"@stll/folio-core": patch
"@stll/folio-react": patch
---

Typing no longer makes the browser rebuild paint data for the whole hidden editor twice per keystroke: the input anchor that follows the caret removed its transform to measure and set it again, and now measures in place and writes only when its offset changes.
