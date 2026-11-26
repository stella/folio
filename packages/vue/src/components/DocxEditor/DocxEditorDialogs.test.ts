import { compileScript, parse } from "@vue/compiler-sfc";
import { plugin } from "bun";
import { afterAll, expect, test } from "bun:test";

import { acquireHarnessDom } from "../../../../../test/harnessDom";

const releaseDom = acquireHarnessDom();
afterAll(releaseDom);

// Compile the real SFCs, so the editor's dialog cluster loads its dialogs the
// way a bundled build does.
plugin({
  name: "docx-editor-dialogs-sfc-test",
  setup(build) {
    build.onLoad(
      { filter: /\/packages\/vue\/src\/.*\.vue$/ },
      async ({ path }) => {
        const { descriptor } = parse(await Bun.file(path).text(), {
          filename: path,
        });
        const compiled = compileScript(descriptor, {
          id: path,
          inlineTemplate: true,
        });
        return { contents: compiled.content, loader: "ts" };
      },
    );
  },
});

const { createApp, defineComponent, h, nextTick, reactive } =
  await import("vue");
const { default: DocxEditorDialogs } = await import("./DocxEditorDialogs.vue");

const settle = async (until: () => boolean) => {
  for (let attempt = 0; attempt < 50 && !until(); attempt++) {
    await new Promise((resolve) => setTimeout(resolve, 0));
    await nextTick();
  }
};

const symbolSearch = () =>
  document.body.querySelector<HTMLInputElement>(
    ".symbol-dialog input.symbol-search",
  );

const recentCells = () =>
  Array.from(
    document.body.querySelectorAll(".symbol-recent .symbol-cell"),
    (recent) => recent.textContent?.trim(),
  );

test("an editor dialog loads on first open, sets itself up, and unmounts on close", async () => {
  const container = document.createElement("div");
  document.body.append(container);
  const state = reactive({ showInsertSymbol: false });
  const inserted: (string | undefined)[] = [];
  const app = createApp(
    defineComponent({
      setup() {
        return () =>
          h(DocxEditorDialogs, {
            search: {} as never,
            view: null,
            bookmarks: [],
            selectedImagePmPos: null,
            sectionProperties: null,
            showFindReplace: false,
            showHyperlink: false,
            showInsertSymbol: state.showInsertSymbol,
            showImageProperties: false,
            showPageSetup: false,
            showTableProperties: false,
            showWatermark: false,
            showNumberingValue: false,
            currentWatermark: undefined,
            tableProperties: {} as never,
            "onUpdate:showInsertSymbol": (value: boolean) => {
              state.showInsertSymbol = value;
            },
            onInsertSymbol: (symbol: string) => inserted.push(symbol),
          });
      },
    }),
  );
  app.mount(container);
  try {
    await settle(() => false);
    expect(symbolSearch()).toBeNull();

    state.showInsertSymbol = true;
    await settle(() => document.activeElement === symbolSearch());
    expect(symbolSearch()).not.toBeNull();
    expect(document.activeElement).toBe(symbolSearch());
    expect(recentCells()).toEqual([]);

    const cell = document.body.querySelector<HTMLButtonElement>(
      ".symbol-grid .symbol-cell",
    );
    cell?.dispatchEvent(new MouseEvent("dblclick", { bubbles: true }));
    expect(inserted).toEqual([cell?.textContent?.trim()]);

    state.showInsertSymbol = false;
    await settle(() => symbolSearch() === null);
    expect(symbolSearch()).toBeNull();

    // Reopening mounts a new dialog that still lists the recent symbol.
    state.showInsertSymbol = true;
    await settle(() => symbolSearch() !== null);
    expect(recentCells()).toEqual(inserted);
  } finally {
    app.unmount();
    container.remove();
  }
});
