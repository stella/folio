import { useState } from "react";
import { useTranslations } from "use-intl";

import { SYMBOL_CATEGORIES, filterSymbols, type SymbolSearchEntry } from "@stll/folio-core/symbols";

import { useFolioUI } from "../../ui/folio-ui";
import {
  DIALOG_BACKDROP_CLASS,
  DIALOG_BODY_CLASS,
  DIALOG_FOOTER_CLASS,
  DIALOG_INPUT_CLASS,
  DIALOG_POPUP_CLASS,
  DIALOG_PRIMARY_BUTTON_CLASS,
  DIALOG_SECONDARY_BUTTON_CLASS,
  DIALOG_TITLE_CLASS,
  useCloseOnDialogOpenChange,
} from "./dialogChrome";

export type InsertSymbolDialogProps = {
  isOpen: boolean;
  onClose: () => void;
  onInsert: (symbol: string) => void;
};

const RECENT_LIMIT = 10;

/** Format a character as its `U+XXXX` code point label. */
function codePointLabel(char: string): string {
  return `U+${(char.codePointAt(0) ?? 0).toString(16).toUpperCase().padStart(4, "0")}`;
}

export function InsertSymbolDialog({ isOpen, onClose, onInsert }: InsertSymbolDialogProps) {
  const {
    Root: Dialog,
    Portal: DialogPortal,
    Backdrop: DialogBackdrop,
    Popup: DialogPopup,
    Title: DialogTitle,
  } = useFolioUI().Dialog;
  const handleOpenChange = useCloseOnDialogOpenChange(onClose);
  const t = useTranslations("folio");
  const [search, setSearch] = useState("");
  const [activeCategory, setActiveCategory] = useState("Common");
  const [selectedSymbol, setSelectedSymbol] = useState("");
  const [recentSymbols, setRecentSymbols] = useState<string[]>([]);

  const activeSymbols: SymbolSearchEntry[] = search
    ? filterSymbols(search)
    : (SYMBOL_CATEGORIES.find((category) => category.name === activeCategory)?.symbols.map(
        (symbol) => Object.assign({}, symbol, { category: activeCategory }),
      ) ?? []);

  const insertSymbol = (symbol: string) => {
    if (!symbol) {
      return;
    }
    setRecentSymbols((prev) =>
      [symbol, ...prev.filter((s) => s !== symbol)].slice(0, RECENT_LIMIT),
    );
    onInsert(symbol);
  };

  return (
    <Dialog open={isOpen} onOpenChange={handleOpenChange}>
      <DialogPortal>
        <DialogBackdrop className={DIALOG_BACKDROP_CLASS} />
        <DialogPopup className={DIALOG_POPUP_CLASS}>
          <DialogTitle className={DIALOG_TITLE_CLASS}>
            {t("dialogs.insertSymbol.title")}
          </DialogTitle>

          <div className={DIALOG_BODY_CLASS}>
            <input
              autoFocus
              value={search}
              onChange={(event) => setSearch(event.target.value)}
              onKeyDown={(event) => {
                if (event.key === "Escape") {
                  onClose();
                }
              }}
              placeholder={t("dialogs.insertSymbol.searchPlaceholder")}
              className={`${DIALOG_INPUT_CLASS} focus:border-primary w-full`}
            />

            {!search && (
              <div className="flex flex-wrap gap-0.5">
                {SYMBOL_CATEGORIES.map((category) => {
                  const active = activeCategory === category.name;
                  return (
                    <button
                      key={category.name}
                      type="button"
                      onClick={() => setActiveCategory(category.name)}
                      className={`rounded px-2 py-1 text-[11px] ${
                        active
                          ? "bg-accent text-accent-foreground font-semibold"
                          : "text-muted-foreground hover:bg-muted"
                      }`}
                    >
                      {t(category.nameKey)}
                    </button>
                  );
                })}
              </div>
            )}

            <div className="grid max-h-60 grid-cols-[repeat(auto-fill,minmax(2.25rem,1fr))] gap-0.5 overflow-y-auto rounded border p-1">
              {activeSymbols.map((symbol) => {
                const selected = selectedSymbol === symbol.char;
                return (
                  <button
                    key={symbol.char}
                    type="button"
                    title={symbol.name}
                    onClick={() => setSelectedSymbol(symbol.char)}
                    onDoubleClick={() => insertSymbol(symbol.char)}
                    className={`flex h-9 w-9 items-center justify-center rounded border text-lg ${
                      selected
                        ? "border-primary bg-primary/20"
                        : "hover:bg-muted border-transparent"
                    }`}
                  >
                    {symbol.char}
                  </button>
                );
              })}
              {activeSymbols.length === 0 && (
                <div className="text-muted-foreground col-span-full p-6 text-center text-[13px]">
                  {search
                    ? t("dialogs.insertSymbol.noResults", { query: search })
                    : t("dialogs.insertSymbol.noResultsEmpty")}
                </div>
              )}
            </div>

            {selectedSymbol && (
              <div className="bg-muted flex items-center gap-3 rounded px-3 py-2">
                <span className="text-3xl leading-none">{selectedSymbol}</span>
                <span className="text-muted-foreground font-mono text-xs">
                  {codePointLabel(selectedSymbol)}
                </span>
              </div>
            )}

            {recentSymbols.length > 0 && !search && (
              <div className="flex flex-wrap items-center gap-1">
                <span className="text-muted-foreground mr-1 text-[11px]">
                  {t("dialogs.insertSymbol.recent")}
                </span>
                {recentSymbols.map((symbol) => (
                  <button
                    key={symbol}
                    type="button"
                    title={codePointLabel(symbol)}
                    onClick={() => setSelectedSymbol(symbol)}
                    onDoubleClick={() => insertSymbol(symbol)}
                    className="hover:bg-muted flex h-7 w-7 items-center justify-center rounded border border-transparent text-sm"
                  >
                    {symbol}
                  </button>
                ))}
              </div>
            )}
          </div>

          <div className={DIALOG_FOOTER_CLASS}>
            <button type="button" className={DIALOG_SECONDARY_BUTTON_CLASS} onClick={onClose}>
              {t("common.cancel")}
            </button>
            <button
              type="button"
              className={DIALOG_PRIMARY_BUTTON_CLASS}
              disabled={!selectedSymbol}
              onClick={() => insertSymbol(selectedSymbol)}
            >
              {t("common.insert")}
            </button>
          </div>
        </DialogPopup>
      </DialogPortal>
    </Dialog>
  );
}
