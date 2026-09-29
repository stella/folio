/** A small seeded generator (mulberry32) so a fuzz run replays from its seed. */
export type Random = {
  /** A float in [0, 1). */
  next: () => number;
  /** An integer in [0, bound). */
  int: (bound: number) => number;
  pick: <T>(items: readonly T[]) => T;
  chance: (probability: number) => boolean;
  /**
   * The generator's position, without advancing it: `createRandom(state())`
   * draws exactly what this one draws next. A flow file records it per step.
   */
  state: () => number;
};

export const createRandom = (seed: number): Random => {
  let state = seed >>> 0;
  const next = (): number => {
    state = (state + 0x6d2b79f5) >>> 0;
    let value = state;
    value = Math.imul(value ^ (value >>> 15), value | 1);
    value ^= value + Math.imul(value ^ (value >>> 7), value | 61);
    return ((value ^ (value >>> 14)) >>> 0) / 4_294_967_296;
  };
  const int = (bound: number): number => Math.floor(next() * bound);
  return {
    next,
    int,
    pick: <T>(items: readonly T[]): T => {
      const item = items[int(items.length)];
      if (item === undefined) {
        throw new Error("pick from an empty list");
      }
      return item;
    },
    chance: (probability: number): boolean => next() < probability,
    state: (): number => state,
  };
};

const WORDS = [
  "agreement",
  "supplier",
  "buyer",
  "delivery",
  "payment",
  "notice",
  "term",
  "warranty",
  "goods",
  "invoice",
  "party",
  "schedule",
  "clause",
  "period",
  "written",
];

/** A short synthetic sentence. */
export const sentence = (random: Random): string => {
  const length = 3 + random.int(6);
  const words = Array.from({ length }, () => random.pick(WORDS));
  const first = words[0] ?? "the";
  return `${first.charAt(0).toUpperCase()}${first.slice(1)} ${words.slice(1).join(" ")}.`;
};
