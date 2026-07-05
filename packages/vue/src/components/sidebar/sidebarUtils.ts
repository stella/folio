/**
 * Single shared avatar style factory + Vue-only re-exports of the data
 * helpers that live in the framework-neutral util layer. The pure helpers
 * (getCommentText etc.) are shared with the React adapter, so we re-export
 * them from folio's util surface instead of forking.
 */
import type { CSSProperties } from "vue";

export {
  getCommentText,
  formatDate,
  getInitials,
  getAvatarColor,
  truncateText,
  type TrackedChangeEntry,
} from "../../utils/comments";

import { getAvatarColor } from "../../utils/comments";

/** Inline style for an avatar bubble — mirrors React's avatarStyle(). */
export function avatarStyle(name: string, size: 32 | 28 = 32): CSSProperties {
  return {
    width: size + "px",
    height: size + "px",
    borderRadius: "50%",
    backgroundColor: getAvatarColor(name),
    color: "#fff",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    fontSize: size === 32 ? "13px" : "11px",
    fontWeight: 500,
    flexShrink: 0,
  };
}
