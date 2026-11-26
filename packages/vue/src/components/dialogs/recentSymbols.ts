import { ref } from "vue";

const RECENT_LIMIT = 10;

/**
 * Symbols inserted through the Insert Symbol dialog, newest first. The editor
 * mounts the dialog only while it is open, so the list lives here and persists
 * across openings.
 */
export const recentSymbols = ref<readonly string[]>([]);

export const rememberRecentSymbol = (symbol: string): void => {
  recentSymbols.value = [
    symbol,
    ...recentSymbols.value.filter((recent) => recent !== symbol),
  ].slice(0, RECENT_LIMIT);
};
