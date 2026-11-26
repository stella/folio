const RECENT_LIMIT = 10;

// Symbols inserted through the Insert Symbol dialog, newest first. The editor
// mounts the dialog only while it is open, so the list lives here and persists
// across openings.
let recentSymbols: readonly string[] = [];

export const getRecentSymbols = (): readonly string[] => recentSymbols;

export const rememberRecentSymbol = (symbol: string): readonly string[] => {
  recentSymbols = [
    symbol,
    ...recentSymbols.filter((recent) => recent !== symbol),
  ].slice(0, RECENT_LIMIT);
  return recentSymbols;
};
