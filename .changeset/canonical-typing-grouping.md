---
"@stll/docx-core": patch
---

In the experimental canonical session, consecutive tracked typing and deletion by one author now read as one change, as in the default editor. A word typed while suggesting saves as one insertion instead of one per character, and a run of Backspace or Delete saves as one deletion. Text typed beside the author's own pending insertion goes into it, and a deletion beside the author's own pending deletion continues it, also when that change came with the file. Another author's change is never continued. Undo still takes back each edit, and rejecting gives the original text back with its runs whole.
