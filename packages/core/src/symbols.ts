/**
 * Symbol catalog for the "Insert Symbol" dialog.
 *
 * Framework-neutral data + filtering shared by every folio adapter (React,
 * Vue, ...), so the two never drift on which symbols they offer. Category
 * display names are i18n keys (`dialogs.insertSymbol.categories.*`) resolved by
 * each adapter's own translator; the `char`/`name` pairs are the raw catalog.
 */

export type SymbolEntry = {
  /** The character to insert. */
  char: string;
  /** English descriptive name, used for search matching. */
  name: string;
};

/** i18n message keys for the symbol-category display labels. */
export type SymbolCategoryNameKey =
  | "dialogs.insertSymbol.categories.common"
  | "dialogs.insertSymbol.categories.arrows"
  | "dialogs.insertSymbol.categories.math"
  | "dialogs.insertSymbol.categories.greek"
  | "dialogs.insertSymbol.categories.currency"
  | "dialogs.insertSymbol.categories.shapes";

export type SymbolCategory = {
  /** Stable category id (also the default active tab). */
  name: string;
  /** i18n key for the category's display label. */
  nameKey: SymbolCategoryNameKey;
  symbols: SymbolEntry[];
};

/** A catalog entry tagged with the category it belongs to (search results). */
export type SymbolSearchEntry = SymbolEntry & { category: string };

export const SYMBOL_CATEGORIES: SymbolCategory[] = [
  {
    name: "Common",
    nameKey: "dialogs.insertSymbol.categories.common",
    symbols: [
      { char: "©", name: "Copyright" },
      { char: "®", name: "Registered" },
      { char: "™", name: "Trademark" },
      { char: "•", name: "Bullet" },
      { char: "…", name: "Ellipsis" },
      { char: "—", name: "Em dash" },
      { char: "–", name: "En dash" },
      { char: "±", name: "Plus-minus" },
      { char: "×", name: "Multiply" },
      { char: "÷", name: "Divide" },
      { char: "≠", name: "Not equal" },
      { char: "≈", name: "Approximately" },
      { char: "≤", name: "Less or equal" },
      { char: "≥", name: "Greater or equal" },
      { char: "°", name: "Degree" },
      { char: "µ", name: "Micro" },
      { char: "¶", name: "Pilcrow" },
      { char: "§", name: "Section" },
      { char: "†", name: "Dagger" },
      { char: "‡", name: "Double dagger" },
      { char: "¿", name: "Inverted question" },
      { char: "¡", name: "Inverted exclamation" },
      { char: "‰", name: "Per mille" },
      { char: "∞", name: "Infinity" },
    ],
  },
  {
    name: "Arrows",
    nameKey: "dialogs.insertSymbol.categories.arrows",
    symbols: [
      { char: "←", name: "Left" },
      { char: "→", name: "Right" },
      { char: "↑", name: "Up" },
      { char: "↓", name: "Down" },
      { char: "↔", name: "Left-right" },
      { char: "↕", name: "Up-down" },
      { char: "⇐", name: "Double left" },
      { char: "⇒", name: "Double right" },
      { char: "⇑", name: "Double up" },
      { char: "⇓", name: "Double down" },
      { char: "⇔", name: "Double left-right" },
      { char: "➡", name: "Heavy right" },
      { char: "↩", name: "Return" },
      { char: "↪", name: "Curved right" },
      { char: "↻", name: "Clockwise" },
      { char: "↺", name: "Counter-clockwise" },
    ],
  },
  {
    name: "Math",
    nameKey: "dialogs.insertSymbol.categories.math",
    symbols: [
      { char: "∑", name: "Summation" },
      { char: "∏", name: "Product" },
      { char: "∫", name: "Integral" },
      { char: "√", name: "Square root" },
      { char: "∂", name: "Partial diff" },
      { char: "∇", name: "Nabla" },
      { char: "∈", name: "Element of" },
      { char: "∉", name: "Not element" },
      { char: "⊂", name: "Subset" },
      { char: "⊃", name: "Superset" },
      { char: "∪", name: "Union" },
      { char: "∩", name: "Intersection" },
      { char: "∧", name: "And" },
      { char: "∨", name: "Or" },
      { char: "¬", name: "Not" },
      { char: "∀", name: "For all" },
      { char: "∃", name: "Exists" },
      { char: "∅", name: "Empty set" },
      { char: "∝", name: "Proportional" },
      { char: "∠", name: "Angle" },
    ],
  },
  {
    name: "Greek",
    nameKey: "dialogs.insertSymbol.categories.greek",
    symbols: [
      { char: "α", name: "alpha" },
      { char: "β", name: "beta" },
      { char: "γ", name: "gamma" },
      { char: "δ", name: "delta" },
      { char: "ε", name: "epsilon" },
      { char: "ζ", name: "zeta" },
      { char: "η", name: "eta" },
      { char: "θ", name: "theta" },
      { char: "λ", name: "lambda" },
      { char: "μ", name: "mu" },
      { char: "π", name: "pi" },
      { char: "ρ", name: "rho" },
      { char: "σ", name: "sigma" },
      { char: "τ", name: "tau" },
      { char: "φ", name: "phi" },
      { char: "ψ", name: "psi" },
      { char: "ω", name: "omega" },
      { char: "Δ", name: "Delta" },
      { char: "Σ", name: "Sigma" },
      { char: "Ω", name: "Omega" },
      { char: "Π", name: "Pi" },
      { char: "Φ", name: "Phi" },
      { char: "Ψ", name: "Psi" },
      { char: "Θ", name: "Theta" },
    ],
  },
  {
    name: "Currency",
    nameKey: "dialogs.insertSymbol.categories.currency",
    symbols: [
      { char: "$", name: "Dollar" },
      { char: "€", name: "Euro" },
      { char: "£", name: "Pound" },
      { char: "¥", name: "Yen" },
      { char: "₹", name: "Rupee" },
      { char: "₽", name: "Ruble" },
      { char: "₩", name: "Won" },
      { char: "₿", name: "Bitcoin" },
      { char: "¢", name: "Cent" },
      { char: "₫", name: "Dong" },
      { char: "₺", name: "Lira" },
      { char: "₴", name: "Hryvnia" },
    ],
  },
  {
    name: "Shapes",
    nameKey: "dialogs.insertSymbol.categories.shapes",
    symbols: [
      { char: "■", name: "Black square" },
      { char: "□", name: "White square" },
      { char: "▲", name: "Up triangle" },
      { char: "▼", name: "Down triangle" },
      { char: "●", name: "Black circle" },
      { char: "○", name: "White circle" },
      { char: "◆", name: "Black diamond" },
      { char: "◇", name: "White diamond" },
      { char: "★", name: "Black star" },
      { char: "☆", name: "White star" },
      { char: "♠", name: "Spade" },
      { char: "♥", name: "Heart" },
      { char: "♦", name: "Diamond" },
      { char: "♣", name: "Club" },
      { char: "✓", name: "Check mark" },
      { char: "✗", name: "Ballot X" },
      { char: "✦", name: "Four pointed star" },
      { char: "◌", name: "Dotted circle" },
    ],
  },
];

/** Every catalog symbol, flattened and tagged with its category. */
export const ALL_SYMBOLS: SymbolSearchEntry[] = SYMBOL_CATEGORIES.flatMap((category) =>
  category.symbols.map((symbol) => ({ ...symbol, category: category.name })),
);

/**
 * Symbols whose descriptive name contains `query` (case-insensitive), or whose
 * character is exactly `query`. An empty/whitespace query returns `[]` so
 * callers fall back to the active category.
 */
export function filterSymbols(query: string): SymbolSearchEntry[] {
  const q = query.trim().toLowerCase();
  if (!q) {
    return [];
  }
  return ALL_SYMBOLS.filter((symbol) => symbol.name.toLowerCase().includes(q) || symbol.char === q);
}
