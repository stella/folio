import {
  ensureLiveView,
  expect,
  forEachAdapter,
  openCanonicalEditor,
  openEditor,
} from "./parity-fixture";

const describeEvents = () => {
  const events: string[] = [];
  const describe = (node: unknown): string => {
    if (!(node instanceof Element)) return String(node);
    return `${node.nodeName}.${String(node.className).slice(0, 40)}`;
  };
  for (const type of ["keydown", "copy", "cut", "paste"]) {
    document.addEventListener(
      type,
      (event) => {
        const selection = document.getSelection();
        const anchor = selection?.anchorNode ?? null;
        const anchorElement =
          anchor instanceof Element ? anchor : (anchor?.parentElement ?? null);
        events.push(
          [
            type,
            "key" in event ? String(event.key) : "",
            `target=${describe(event.target)}`,
            `active=${describe(document.activeElement)}`,
            `selection=${String(selection?.toString().length ?? -1)}`,
            `anchorInPM=${String(anchorElement?.closest(".ProseMirror") !== null)}`,
            `prevented=${String(event.defaultPrevented)}`,
          ].join(" "),
        );
      },
      true,
    );
  }
  Object.assign(window, { __clipboardDiagnostic: events });
};

for (const [engine, open] of [
  ["legacy", openEditor],
  ["canonical", openCanonicalEditor],
] as const) {
  forEachAdapter(
    `clipboard diagnostic ${engine}`,
    async (adapter, { page }) => {
      await page.addInitScript(describeEvents);
      await open(page, adapter);
      await ensureLiveView(page);
      await page.locator(".layout-paragraph").first().click();
      await page.keyboard.press("Home");
      await page.keyboard.press("Shift+End");
      await page.keyboard.press("ControlOrMeta+c");
      await page.waitForTimeout(500);
      const diagnostic = await page.evaluate(() =>
        JSON.stringify(
          (window as unknown as { __clipboardDiagnostic?: string[] })
            .__clipboardDiagnostic,
        ),
      );
      const count = await page.evaluate(
        () => window.__folioParity?.getClipboardCallbackCount("copy") ?? -1,
      );
      expect(`${engine} count=${String(count)} ${diagnostic}`).toBe("");
    },
  );
}
