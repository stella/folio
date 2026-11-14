---
"@stll/folio-core": patch
"@stll/folio-react": patch
"@stll/folio-vue": patch
---

Dragging the ruler's left and right indents and double-clicking a tab stop to remove it work in a canonical session in both adapters: the edit runs through `applyRulerParagraphEdit` (a new core command owner) and the view's command owner, so the session journals it as paragraph formatting. The Vue ruler's indent handles carry translated labels.
