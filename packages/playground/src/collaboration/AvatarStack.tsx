import { sliceText } from "@stll/folio-core/utils/sliceText";

import type { CollaborativeUser } from "./useCollaboration";

type AvatarStackProps = {
  users: CollaborativeUser[];
  max?: number;
};

const initials = (name: string): string => {
  const parts = name.trim().split(/\s+/);
  if (parts.length === 1) {
    return sliceText({ text: parts[0]!, units: "graphemes", end: 2 });
  }
  return parts
    .filter((_, index) => index === 0 || index === parts.length - 1)
    .map((text) => sliceText({ text, units: "graphemes", end: 1 }))
    .join("");
};

export function AvatarStack({ users, max = 5 }: AvatarStackProps) {
  if (users.length === 0) {
    return null;
  }

  const sorted = [...users].sort(
    (a, b) => Number(b.isLocal) - Number(a.isLocal),
  );
  const visible = sorted.slice(0, max);
  const overflow = sorted.length - visible.length;

  return (
    <div
      className="pg-avatar-stack"
      aria-label={`${users.length} active collaborator(s)`}
    >
      {visible.map((user) => (
        <div
          key={user.clientId}
          className="pg-avatar"
          style={{ background: user.color }}
          title={user.isLocal ? `${user.name} (you)` : user.name}
        >
          {initials(user.name)}
        </div>
      ))}
      {overflow > 0 && (
        <div
          className="pg-avatar pg-avatar--overflow"
          title={`${overflow} more`}
        >
          +{overflow}
        </div>
      )}
    </div>
  );
}
