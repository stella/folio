---
"@stll/folio-core": patch
---

The canonical session holds a content lock on a revision around a whole control as well as one inside it: rejecting a deletion that wraps a `contentLocked` control stays pending, while a new tracked deletion of the whole control is judged as the control's removal.
