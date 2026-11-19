/** The one semantic compiler for canonical sidebar, controlled and public comments. */
import {
  type Err,
  type Ok,
  Result,
  TaggedError,
  type TaggedErrorClass,
} from "better-result";

import {
  allocateCommentAnchorIds,
  applyDocumentOps,
  type CreateCommentOp,
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  formatParaId,
  freshLeasedIds,
  ID_KIND_LIMITS,
  ID_LEASE_KINDS,
  type IdLease,
  packageParagraphIds,
  type UpdateCommentContentOp,
} from "@stll/docx-core/ops";

import { commentThreadParaId } from "../docx/commentThreadKey";
import { threadedCommentIds } from "../docx/serializer/commentSerializer";
import { CANONICAL_GAP } from "../types/canonicalCapabilities";
import type { Comment } from "../types/content";
import type { Document } from "../types/document";
import { canonicalJson } from "../utils/canonicalJson";

const CanonicalCommentErrorBase: TaggedErrorClass<"CanonicalCommentError"> =
  TaggedError("CanonicalCommentError");

export class CanonicalCommentError extends CanonicalCommentErrorBase<{
  message: string;
  gap: typeof CANONICAL_GAP.comments;
}> {}

const refuse = (message: string) =>
  Result.err(
    new CanonicalCommentError({ message, gap: CANONICAL_GAP.comments }),
  );

export type CanonicalCommentCommand =
  | {
      type: "create";
      comment: CreateCommentOp["comment"];
      anchor: CreateCommentOp["anchor"];
    }
  | { type: "update"; id: number; content: Comment["content"] }
  | { type: "resolve"; id: number; status: "open" | "resolved" }
  | { type: "delete"; id: number }
  | { type: "replace"; comments: readonly Comment[] };

type CompileCanonicalCommentsOptions = {
  document: Document;
  command: CanonicalCommentCommand;
  /** The ids the author may mint, when it holds a lease. */
  idLease?: IdLease | undefined;
};

export const compileCanonicalComments = ({
  document,
  command,
  idLease,
}: CompileCanonicalCommentsOptions):
  | Err<never, CanonicalCommentError>
  | Ok<
      {
        ops: DocumentOp[];
        document: Document;
        commentId: number | undefined;
      },
      never
    > => {
  const current = document.package.document.comments ?? [];
  const byId = new Map(current.map((comment) => [comment.id, comment]));
  const ops: DocumentOp[] = [];
  let commentId: number | undefined;
  switch (command.type) {
    case "create": {
      ops.push({
        type: DOCUMENT_OP_TYPES.CREATE_COMMENT,
        comment: command.comment,
        anchor: command.anchor,
      });
      commentId = command.comment.id;
      break;
    }
    case "update":
      ops.push({
        type: DOCUMENT_OP_TYPES.UPDATE_COMMENT_CONTENT,
        id: command.id,
        content: command.content,
      });
      break;
    case "resolve":
      ops.push({
        type: DOCUMENT_OP_TYPES.SET_COMMENT_RESOLUTION,
        id: command.id,
        status: command.status,
      });
      break;
    case "delete": {
      const comment = byId.get(command.id);
      if (!comment) return refuse("The comment no longer exists.");
      ops.push({
        type: DOCUMENT_OP_TYPES.DELETE_COMMENT,
        id: command.id,
        scope:
          comment.parentId !== undefined && byId.has(comment.parentId)
            ? "reply"
            : "thread",
      });
      break;
    }
    case "replace": {
      const nextById = new Map(
        command.comments.map((comment) => [comment.id, comment]),
      );
      if (nextById.size !== command.comments.length)
        return refuse("Comment ids must be unique.");
      for (const comment of current) {
        if (
          nextById.has(comment.id) &&
          comment.parentId !== undefined &&
          byId.has(comment.parentId) &&
          !nextById.has(comment.parentId)
        )
          return refuse("A retained reply requires its parent thread.");

        if (nextById.has(comment.id)) continue;
        if (
          comment.parentId !== undefined &&
          !nextById.has(comment.parentId) &&
          byId.has(comment.parentId)
        )
          continue;
        ops.push({
          type: DOCUMENT_OP_TYPES.DELETE_COMMENT,
          id: comment.id,
          scope:
            comment.parentId !== undefined && byId.has(comment.parentId)
              ? "reply"
              : "thread",
        });
      }
      const pending = command.comments.filter(
        (comment) => !byId.has(comment.id),
      );
      const available = new Set(
        current
          .filter((comment) => nextById.has(comment.id))
          .map((comment) => comment.id),
      );
      while (pending.length > 0) {
        const index = pending.findIndex(
          (comment) =>
            comment.parentId !== undefined && available.has(comment.parentId),
        );
        if (index < 0)
          return refuse(
            "New controlled comments require an existing anchored parent without cycles.",
          );
        const next = pending.splice(index, 1).at(0);
        if (next === undefined || next.parentId === undefined)
          return refuse("The controlled reply is missing its parent.");
        const { parentId, ...comment } = next;
        ops.push({
          type: DOCUMENT_OP_TYPES.CREATE_COMMENT,
          comment,
          anchor: { kind: "reply", parentId },
        });
        available.add(next.id);
      }
      for (const next of command.comments) {
        const previous = byId.get(next.id);
        if (!previous) continue;
        const {
          content: priorContent,
          done: priorDone,
          author: priorAuthor,
          initials: priorInitials,
          date: priorDate,
          ...priorMetadata
        } = previous;
        const {
          content: nextContent,
          done: nextDone,
          author: nextAuthor,
          initials: nextInitials,
          date: nextDate,
          ...nextMetadata
        } = next;
        if (canonicalJson(priorMetadata) !== canonicalJson(nextMetadata))
          return refuse(
            "Controlled comment identity and opaque metadata cannot be replaced.",
          );
        const patch = {
          ...(priorAuthor === nextAuthor ? {} : { author: nextAuthor }),
          ...(priorInitials === nextInitials
            ? {}
            : {
                initials: Object.hasOwn(next, "initials") ? nextInitials : null,
              }),
          ...(priorDate === nextDate
            ? {}
            : { date: Object.hasOwn(next, "date") ? nextDate : null }),
        } satisfies NonNullable<UpdateCommentContentOp["patch"]>;
        if (
          canonicalJson(priorContent) !== canonicalJson(nextContent) ||
          Object.keys(patch).length > 0
        )
          ops.push({
            type: DOCUMENT_OP_TYPES.UPDATE_COMMENT_CONTENT,
            id: next.id,
            content: nextContent,
            patch,
          });
        if (priorDone !== nextDone) {
          if (nextDone === undefined)
            return refuse(
              "Resolution changes must name an open or resolved state.",
            );
          ops.push({
            type: DOCUMENT_OP_TYPES.SET_COMMENT_RESOLUTION,
            id: next.id,
            status: nextDone ? "resolved" : "open",
          });
        }
      }
      break;
    }
    default: {
      const unreachable: never = command;
      return unreachable;
    }
  }
  const compiled: DocumentOp[] = [];
  let draft = document;
  for (const op of ops) {
    const materialized = (() => {
      if (op.type !== DOCUMENT_OP_TYPES.CREATE_COMMENT) return Result.ok(op);
      const allocation = allocateCommentAnchorIds(draft, op, idLease);
      if (allocation.isErr()) return Result.err(allocation.error);
      return Result.ok({ ...op, newIds: allocation.value });
    })();
    if (materialized.isErr()) return refuse(materialized.error.message);
    const applied = applyDocumentOps(draft, [materialized.value]);
    if (applied.isErr()) return refuse(applied.error.message);
    compiled.push(materialized.value);
    draft = applied.value.document;
  }
  // A reply, its thread root and a resolved comment are linked through the
  // paragraph id of their last paragraph; the save mints one where the model
  // has none. The command mints it instead, so the model is what saves.
  const threaded = threadedCommentIds(draft.package.document.comments ?? []);
  for (const comment of draft.package.document.comments ?? []) {
    const content = comment.content ?? [];
    const last = content.at(-1);
    if (
      !threaded.has(comment.id) ||
      last === undefined ||
      commentThreadParaId(content.map(({ paraId }) => paraId)) !== undefined
    )
      continue;
    const paraId = canonicalCommentBody(draft, "", idLease).at(0)?.paraId;
    if (paraId === undefined) return refuse("No paragraph id is free.");
    const op: DocumentOp = {
      type: DOCUMENT_OP_TYPES.UPDATE_COMMENT_CONTENT,
      id: comment.id,
      content: [...content.slice(0, -1), { ...last, paraId }],
    };
    const applied = applyDocumentOps(draft, [op]);
    if (applied.isErr()) return refuse(applied.error.message);
    compiled.push(op);
    draft = applied.value.document;
  }
  return Result.ok({ ops: compiled, document: draft, commentId });
};

export const canonicalCommentBody = (
  document: Document,
  text: string,
  lease?: IdLease,
): Comment["content"] => {
  const occupied = new Set(
    packageParagraphIds(document.package).map((id) => id.toUpperCase()),
  );
  // An exhausted lease gives an id no operation accepts: the edit is refused.
  const id =
    freshLeasedIds({
      lease,
      kind: ID_LEASE_KINDS.PARAGRAPH,
      count: 1,
      taken: (candidate) => occupied.has(formatParaId(candidate)),
    }).at(0) ?? ID_KIND_LIMITS.paragraph;
  return [
    {
      type: "paragraph",
      paraId: formatParaId(id),
      content: [{ type: "run", content: [{ type: "text", text }] }],
    },
  ];
};
