---
"@stll/folio-core": patch
"@stll/folio-react": patch
"@stll/folio-vue": patch
---

Comment edits are part of the editor's undo history. Deleting a thread (with its replies and anchors), resolving or reopening it, replying and adding a comment can each be undone and redone, in React and Vue, and the save writes the restored threads. Accepting or rejecting a change that removes everything a comment covers removes the comment with it, in the same undo step, and each accept or reject is an undo step of its own. Comments that document operations add join the editor's threads in the operation's own transaction, so undoing the operation removes them too. With a controlled `comments` prop, undo and redo report the restored list through `onCommentsChange`. New in core: `@stll/folio-core/prosemirror/commentRecords`.
