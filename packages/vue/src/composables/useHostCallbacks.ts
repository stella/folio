/**
 * Bridges the host-facing `onComment*` and `onEditorViewReady` props into the
 * editor internals, keeping the host SFC lean.
 *
 * PORT-BLOCKED (missing `useCommentManagement` + comment props): upstream's
 * primary output is a `CommentCallbacks` bundle
 * (`onCommentAdd`/`onCommentResolve`/`onCommentDelete`/`onCommentReply`/
 * `onCommentsChange`) whose type is exported by the sibling
 * `./useCommentManagement` composable — not yet ported to our fork. It also
 * reads four host props (`onCommentAdd`, `onCommentResolve`, `onCommentDelete`,
 * `onCommentReply`) that do not exist on our `DocxEditorProps` (which currently
 * exposes only `onCommentsChange` and `onEditorViewReady`). Both the
 * `CommentCallbacks` contract and the props it wires are absent, so the
 * comment-bridge cannot be built without fabricating an API.
 *
 * The `onEditorViewReady` watch is portable on its own, but it is a bare side
 * effect (no return value) and not the composable's contract, so this is left
 * as a full block rather than a degenerate stub.
 *
 * Unblock: port `useCommentManagement` (exporting `CommentCallbacks`) and add
 * the four `onComment*` callbacks to `DocxEditorProps`, then restore the
 * upstream body verbatim.
 */

export {};
