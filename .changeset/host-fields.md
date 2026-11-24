---
"@stll/folio-core": minor
"@stll/folio-react": minor
"@stll/folio-vue": minor
---

Hosts can insert fields and derive what they show, in either session. `DocxEditorRef.insertField({ instruction, displayText, locked })` puts a `w:fldSimple` in place of the selection: one undo step, a tracked insertion when suggesting. The `derivedFields` prop names the fields a host derives (`owns(instruction)`) and what each shows (`displayTexts(instructions)`, in document order, tracked deletions left out); after each local edit the owned fields are brought up to date in the same undo step, and a peer in a shared room receives the update with the edit.
