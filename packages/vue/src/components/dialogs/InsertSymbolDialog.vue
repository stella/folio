<!--
  Insert Symbol dialog — ported from the upstream docx-editor Vue adapter.
-->
<template>
  <div v-if="isOpen" class="dialog-overlay" @mousedown.self="close">
    <div class="dialog symbol-dialog" @mousedown.stop @keydown.stop>
      <div class="dialog__header">
        <span class="dialog__title">{{ t("dialogs.insertSymbol.title") }}</span>
        <button class="dialog__close" :title="t('common.closeDialog')" @click="close">✕</button>
      </div>
      <div class="dialog__body">
        <!-- Search -->
        <input
          ref="searchRef"
          v-model="search"
          class="symbol-search"
          :placeholder="t('dialogs.insertSymbol.searchPlaceholder')"
          @keydown.escape="close"
        />

        <!-- Category tabs -->
        <div v-if="!search" class="symbol-tabs">
          <button
            v-for="cat in categories"
            :key="cat.name"
            class="symbol-tab"
            :class="{ active: activeCategory === cat.name }"
            @mousedown.prevent="activeCategory = cat.name"
          >
            {{ t(cat.nameKey) }}
          </button>
        </div>

        <!-- Symbol grid -->
        <div class="symbol-grid">
          <button
            v-for="sym in displayedSymbols"
            :key="sym.char"
            class="symbol-cell"
            :class="{ selected: selectedSymbol === sym.char }"
            :title="sym.name"
            @click="selectedSymbol = sym.char"
            @dblclick="insertSymbol(sym.char)"
          >
            {{ sym.char }}
          </button>
          <div v-if="displayedSymbols.length === 0" class="symbol-empty">
            {{
              search
                ? t("dialogs.insertSymbol.noResults", { query: search })
                : t("dialogs.insertSymbol.noResultsEmpty")
            }}
          </div>
        </div>

        <!-- Preview & info -->
        <div v-if="selectedSymbol" class="symbol-preview">
          <span class="symbol-preview__char">{{ selectedSymbol }}</span>
          <span class="symbol-preview__info"
            >U+{{
              selectedSymbol.codePointAt(0)?.toString(16).toUpperCase().padStart(4, "0")
            }}</span
          >
        </div>

        <!-- Recent -->
        <div v-if="recentSymbols.length > 0 && !search" class="symbol-recent">
          <div class="symbol-recent__label">{{ t("dialogs.insertSymbol.recent") }}</div>
          <button
            v-for="s in recentSymbols"
            :key="s"
            class="symbol-cell symbol-cell--small"
            @dblclick="insertSymbol(s)"
            @click="selectedSymbol = s"
          >
            {{ s }}
          </button>
        </div>

        <div class="dialog__actions">
          <button class="dialog__btn" @click="close">{{ t("common.cancel") }}</button>
          <button
            class="dialog__btn dialog__btn--primary"
            @mousedown.prevent="insertSymbol(selectedSymbol)"
            :disabled="!selectedSymbol"
          >
            {{ t("common.insert") }}
          </button>
        </div>
      </div>
    </div>
  </div>
</template>

<script setup lang="ts">
import { ref, computed, watch, nextTick } from "vue";
import { useTranslation } from "../../i18n";
import { SYMBOL_CATEGORIES, filterSymbols, type SymbolSearchEntry } from "@stll/folio-core/symbols";

const { t } = useTranslation();

const props = defineProps<{ isOpen: boolean }>();
const emit = defineEmits<{
  (e: "close"): void;
  (e: "insert", symbol: string): void;
}>();

const searchRef = ref<HTMLInputElement | null>(null);
const search = ref("");
const activeCategory = ref("Common");
const selectedSymbol = ref("");
const recentSymbols = ref<string[]>([]);

const categories = SYMBOL_CATEGORIES;

const displayedSymbols = computed<SymbolSearchEntry[]>(() => {
  if (search.value) {
    return filterSymbols(search.value);
  }
  const cat = categories.find((c) => c.name === activeCategory.value);
  return cat ? cat.symbols.map((symbol) => Object.assign({}, symbol, { category: cat.name })) : [];
});

watch(
  () => props.isOpen,
  async (open) => {
    if (open) {
      await nextTick();
      searchRef.value?.focus();
    }
  },
);

function close() {
  emit("close");
}

function insertSymbol(sym: string) {
  if (!sym) return;
  // Track recent
  recentSymbols.value = [sym, ...recentSymbols.value.filter((s) => s !== sym)].slice(0, 10);
  emit("insert", sym);
}
</script>

<style scoped>
.dialog-overlay {
  position: fixed;
  inset: 0;
  background: var(--doc-overlay);
  z-index: 300;
  display: flex;
  align-items: center;
  justify-content: center;
}
.dialog {
  background: var(--doc-surface);
  border-radius: 8px;
  box-shadow: 0 8px 30px var(--doc-shadow);
  max-width: 90vw;
}
.symbol-dialog {
  width: 480px;
}
.dialog__header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  padding: 12px 16px;
  border-bottom: 1px solid var(--doc-border);
}
.dialog__title {
  font-weight: 600;
  font-size: 14px;
  color: var(--doc-text);
}
.dialog__close {
  border: none;
  background: transparent;
  cursor: pointer;
  font-size: 14px;
  color: var(--doc-text-muted);
  width: 24px;
  height: 24px;
  display: flex;
  align-items: center;
  justify-content: center;
  border-radius: 4px;
}
.dialog__close:hover {
  background: var(--doc-bg-hover);
}
.dialog__body {
  padding: 16px;
}
.dialog__actions {
  display: flex;
  justify-content: flex-end;
  gap: 8px;
  margin-top: 12px;
}
.dialog__btn {
  padding: 6px 16px;
  border: 1px solid var(--doc-border-dark);
  border-radius: 4px;
  cursor: pointer;
  font-size: 13px;
  background: var(--doc-surface);
}
.dialog__btn--primary {
  background: var(--doc-primary);
  color: var(--doc-on-primary);
  border-color: var(--doc-primary);
}
.dialog__btn--primary:hover {
  background: var(--doc-primary);
}
.dialog__btn--primary:disabled {
  opacity: 0.5;
  cursor: not-allowed;
}

.symbol-search {
  width: 100%;
  padding: 6px 10px;
  border: 1px solid var(--doc-border-dark);
  border-radius: 4px;
  font-size: 13px;
  outline: none;
  box-sizing: border-box;
  margin-bottom: 8px;
}
.symbol-search:focus {
  border-color: var(--doc-primary);
}

.symbol-tabs {
  display: flex;
  gap: 2px;
  margin-bottom: 8px;
  flex-wrap: wrap;
}
.symbol-tab {
  padding: 4px 8px;
  border: none;
  border-radius: 4px;
  background: transparent;
  cursor: pointer;
  font-size: 11px;
  color: var(--doc-text-muted);
}
.symbol-tab:hover {
  background: var(--doc-bg-hover);
}
.symbol-tab.active {
  background: var(--doc-accent-bg);
  color: var(--doc-accent);
  font-weight: 600;
}

.symbol-grid {
  display: grid;
  grid-template-columns: repeat(auto-fill, minmax(36px, 1fr));
  gap: 2px;
  max-height: 240px;
  overflow-y: auto;
  border: 1px solid var(--doc-border);
  border-radius: 4px;
  padding: 4px;
}
.symbol-cell {
  width: 36px;
  height: 36px;
  display: flex;
  align-items: center;
  justify-content: center;
  border: 1px solid transparent;
  border-radius: 4px;
  cursor: pointer;
  font-size: 18px;
  background: transparent;
}
.symbol-cell:hover {
  background: var(--doc-bg-hover);
  border-color: var(--doc-border-dark);
}
.symbol-cell.selected {
  background: var(--doc-primary-light);
  border-color: var(--doc-primary);
}
.symbol-cell--small {
  width: 28px;
  height: 28px;
  font-size: 14px;
}
.symbol-empty {
  grid-column: 1/-1;
  text-align: center;
  padding: 24px;
  color: var(--doc-text-subtle);
  font-size: 13px;
}

.symbol-preview {
  display: flex;
  align-items: center;
  gap: 12px;
  margin-top: 8px;
  padding: 8px 12px;
  background: var(--doc-bg);
  border-radius: 4px;
}
.symbol-preview__char {
  font-size: 28px;
}
.symbol-preview__info {
  font-size: 12px;
  color: var(--doc-text-muted);
  font-family: monospace;
}

.symbol-recent {
  display: flex;
  align-items: center;
  gap: 4px;
  margin-top: 8px;
  flex-wrap: wrap;
}
.symbol-recent__label {
  font-size: 11px;
  color: var(--doc-text-subtle);
  margin-right: 4px;
}
</style>
