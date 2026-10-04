import type { OpStory } from "@stll/docx-core/ops";
import type { Comment } from "./content";
import type { CanonicalGap } from "./canonicalCapabilities";

/** Requests name their anchor; comment metadata never creates an unanchored root. */
export type CanonicalCommentRequest =
  | {
      type: "create";
      text: string;
      author: string;
      date?: string;
      anchor:
        | { kind: "selection"; from: number; to: number; story: OpStory }
        | { kind: "reply"; parentId: number }
        | { kind: "revision"; story: OpStory; revisionId: number };
    }
  | { type: "update"; id: number; content: Comment["content"] }
  | { type: "resolve"; id: number; status: "open" | "resolved" }
  | { type: "delete"; id: number }
  | { type: "replace"; comments: readonly Comment[] };

export type CanonicalCommentResult =
  | { status: "applied"; comments: Comment[]; commentId?: number }
  | { status: "refused"; gap: CanonicalGap; message: string };
