// Lightweight per-tab identity for the collaboration demo. Persisted to
// sessionStorage so a refresh keeps the same user; a new tab gets a fresh
// persona for side-by-side multi-user testing.

const NAMES = ["Ada", "Grace", "Linus", "Hedy", "Margaret", "Tim", "Donald", "Barbara", "Alan", "Radia"];

const COLORS = [
  "#ef4444",
  "#f97316",
  "#eab308",
  "#22c55e",
  "#06b6d4",
  "#3b82f6",
  "#8b5cf6",
  "#ec4899",
];

export type CollabUser = {
  name: string;
  color: string;
};

const STORAGE_KEY = "folio-playground-collab-user";

const pick = <T,>(arr: readonly T[]): T => arr[Math.floor(Math.random() * arr.length)]!;

export const loadOrCreateUser = (): CollabUser => {
  try {
    const raw = sessionStorage.getItem(STORAGE_KEY);
    if (raw) {
      return JSON.parse(raw) as CollabUser;
    }
  } catch {
    // ignore corrupt storage
  }
  const user: CollabUser = { name: pick(NAMES), color: pick(COLORS) };
  try {
    sessionStorage.setItem(STORAGE_KEY, JSON.stringify(user));
  } catch {
    // ignore quota errors
  }
  return user;
};

export const getOrCreateRoomFromUrl = (): string => {
  const hash = window.location.hash.replace(/^#/, "");
  if (hash) {
    return hash;
  }
  return `room-${Math.random().toString(36).slice(2, 10)}`;
};
