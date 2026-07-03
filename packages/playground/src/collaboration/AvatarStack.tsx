import type { CollaborativeUser } from "./useCollaboration";

type AvatarStackProps = {
  users: CollaborativeUser[];
  max?: number;
};

const initials = (name: string): string => {
  const parts = name.trim().split(/\s+/);
  if (parts.length === 1) {
    return parts[0]!.slice(0, 2);
  }
  return (parts[0]![0]! + parts.at(-1)![0]!).slice(0, 2);
};

export function AvatarStack({ users, max = 5 }: AvatarStackProps) {
  if (users.length === 0) {
    return null;
  }

  const sorted = [...users].sort((a, b) => Number(b.isLocal) - Number(a.isLocal));
  const visible = sorted.slice(0, max);
  const overflow = sorted.length - visible.length;

  return (
    <div className="pg-avatar-stack" aria-label={`${users.length} active collaborator(s)`}>
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
        <div className="pg-avatar pg-avatar--overflow" title={`${overflow} more`}>
          +{overflow}
        </div>
      )}
    </div>
  );
}
