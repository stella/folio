import { GlobalRegistrator } from "@happy-dom/global-registrator";

GlobalRegistrator.register();

import { afterAll, expect, mock, test } from "bun:test";
import { panic } from "better-result";
import { act, type ReactNode } from "react";
import { createRoot } from "react-dom/client";
import { IntlProvider } from "use-intl";
import { getFolioMessages } from "@stll/folio-core/i18n/messages";
import { FolioUIProvider, type FolioDialog, type FolioInputProps } from "../../ui/folio-ui";
import { ImagePositionDialog } from "./ImagePositionDialog";
import { ImagePropertiesDialog } from "./ImagePropertiesDialog";
import { TablePropertiesDialog } from "./TablePropertiesDialog";
import { SetNumberingValueDialog } from "./SetNumberingValueDialog";
import { PasteSpecialDialog } from "./PasteSpecialDialog";
import { InsertImageDialog } from "./InsertImageDialog";
import { InsertTableDialog } from "./InsertTableDialog";
import { SplitCellDialog } from "./SplitCellDialog";
import { HyperlinkDialog } from "./HyperlinkDialog";
import { WatermarkDialog } from "./WatermarkDialog";
import { PageSetupDialog } from "./PageSetupDialog";
import { FindReplaceDialog } from "./FindReplaceDialog";
import { HyperlinkPopup } from "../ui/HyperlinkPopup";

const previousActEnvironment = Reflect.get(globalThis, "IS_REACT_ACT_ENVIRONMENT");
Reflect.set(globalThis, "IS_REACT_ACT_ENVIRONMENT", true);
afterAll(() => {
  Reflect.set(globalThis, "IS_REACT_ACT_ENVIRONMENT", previousActEnvironment);
  GlobalRegistrator.unregister();
});

// Test the form lifecycle independently of the consumer's dialog portal implementation.
const dialog = {
  Root: ({ children }) => <>{children}</>,
  Portal: ({ children }) => <>{children}</>,
  Backdrop: () => null,
  Popup: ({ children }) => <div>{children}</div>,
  Title: ({ children }) => <h2>{children}</h2>,
  Close: ({ children }) => <button type="button">{children}</button>,
} satisfies FolioDialog;
// The injected input forwards the native input event through the public
// onChange contract, so Happy DOM drives the real dialog event handler.
const NativeEventInput = ({
  onChange,
  size: _size,
  nativeInput: _nativeInput,
  ...props
}: FolioInputProps) => <input {...props} onInput={onChange} />;
const messages = getFolioMessages("en");
const components = { Dialog: dialog, Input: NativeEventInput };
const imagePosition = { horizontal: { posOffset: 10 }, distTop: 20 };
const imageProperties = { alt: "Diagram", borderWidth: 2 };
const tableProperties = { width: 2500, widthType: "dxa", justification: "center" } as const;
const hyperlink = { href: "#intro", displayText: "Introduction" };
const watermark = { kind: "text", text: "DRAFT", opacity: 0.3 } as const;
const firstImage = { alt: "first" };
const secondImage = { alt: "second" };
const externalResult = {
  matches: [0, 10, 20].map((startOffset) => ({
    paragraphIndex: 0,
    contentIndex: 0,
    startOffset,
    endOffset: startOffset + 8,
    text: "selected",
  })),
  totalCount: 3,
  currentIndex: 1,
};
const lateExternalResult = { ...externalResult, currentIndex: 2 };
const popupData = (anchorEl: HTMLAnchorElement, name: string) => ({
  href: `https://${name}.example`,
  displayText: name,
  anchorEl,
});
const noop = () => {};
const wrap = (children: ReactNode) => (
  <IntlProvider locale="en" timeZone="UTC" messages={messages}>
    <FolioUIProvider components={components}>{children}</FolioUIProvider>
  </IntlProvider>
);

const formCases = [
  (isOpen: boolean) => (
    <ImagePositionDialog
      isOpen={isOpen}
      onClose={noop}
      onApply={noop}
      currentData={imagePosition}
    />
  ),
  (isOpen: boolean) => (
    <ImagePropertiesDialog
      isOpen={isOpen}
      onClose={noop}
      onApply={noop}
      currentData={imageProperties}
    />
  ),
  (isOpen: boolean) => (
    <TablePropertiesDialog
      isOpen={isOpen}
      onClose={noop}
      onApply={noop}
      currentProps={tableProperties}
    />
  ),
  (isOpen: boolean) => <SetNumberingValueDialog isOpen={isOpen} onClose={noop} onApply={noop} />,
  (isOpen: boolean) => (
    <PasteSpecialDialog
      isOpen={isOpen}
      onClose={noop}
      onPaste={noop}
      defaultMode="mergeFormatting"
    />
  ),
  (isOpen: boolean) => <InsertImageDialog isOpen={isOpen} onClose={noop} onInsert={noop} />,
  (isOpen: boolean) => (
    <InsertTableDialog
      isOpen={isOpen}
      onClose={noop}
      onInsert={noop}
      defaultRows={5}
      defaultColumns={4}
    />
  ),
  (isOpen: boolean) => (
    <SplitCellDialog
      isOpen={isOpen}
      onClose={noop}
      onSplit={noop}
      defaultRows={2}
      defaultColumns={3}
    />
  ),
  (isOpen: boolean) => (
    <HyperlinkDialog isOpen={isOpen} onClose={noop} onSubmit={noop} currentData={hyperlink} />
  ),
  (isOpen: boolean) => (
    <WatermarkDialog isOpen={isOpen} onClose={noop} onApply={noop} currentWatermark={watermark} />
  ),
  (isOpen: boolean) => <PageSetupDialog isOpen={isOpen} onClose={noop} onApply={noop} />,
];

for (const [index, renderForm] of formCases.entries()) {
  test(`form ${index} initializes immediately and remounts on reopening`, async () => {
    const container = document.createElement("div");
    const root = createRoot(container);
    const values = () =>
      [...container.querySelectorAll("input, select, textarea")].map((field) => {
        if (field instanceof HTMLInputElement) return [field.value, field.checked];
        if (field instanceof HTMLSelectElement || field instanceof HTMLTextAreaElement)
          return field.value;
        return null;
      });
    await act(async () => root.render(wrap(renderForm(true))));
    const originalFields = [...container.querySelectorAll("input, select, textarea")];
    const initialValues = values();
    expect(originalFields.length).toBeGreaterThan(0);
    await act(async () => root.render(wrap(renderForm(false))));
    expect(container.querySelector("input, select, textarea")).toBeNull();
    await act(async () => root.render(wrap(renderForm(true))));
    expect(values()).toEqual(initialValues);
    expect(container.querySelector("input, select, textarea")).not.toBe(originalFields.at(0));
    await act(async () => root.unmount());
  });
}

test("source changes reset editable forms without retaining values from the previous source", async () => {
  const container = document.createElement("div");
  const root = createRoot(container);
  const render = (currentData: { alt: string }) =>
    wrap(<ImagePropertiesDialog isOpen onClose={noop} onApply={noop} currentData={currentData} />);
  await act(async () => root.render(render(firstImage)));
  expect(container.querySelector("textarea")?.value).toBe("first");
  await act(async () => root.render(render(secondImage)));
  expect(container.querySelector("textarea")?.value).toBe("second");
  await act(async () => root.unmount());
});

test("find initializes selected text, searches, clears on close and refreshes external results", async () => {
  const container = document.createElement("div");
  const root = createRoot(container);
  const onFind = mock(() => ({ matches: [], totalCount: 0, currentIndex: -1 }));
  const onClearHighlights = mock(noop);
  const props = {
    onClose: noop,
    onFind,
    onFindNext: () => null,
    onFindPrevious: () => null,
    onReplace: () => false,
    onReplaceAll: () => 0,
    onClearHighlights,
  };
  await act(async () =>
    root.render(wrap(<FindReplaceDialog {...props} isOpen initialSearchText="selected" />)),
  );
  expect(container.querySelector("input")?.value).toBe("selected");
  await act(async () => {
    await new Promise((resolve) => setTimeout(resolve, 140));
  });
  expect(onFind).toHaveBeenCalledWith("selected", { matchCase: false, matchWholeWord: false });
  await act(async () =>
    root.render(
      wrap(
        <FindReplaceDialog
          {...props}
          isOpen
          initialSearchText="selected"
          currentResult={externalResult}
        />,
      ),
    ),
  );
  expect(container.textContent).toContain("2 / 3");
  // Rendering a host result must not schedule another identical search.
  await act(async () => {
    await new Promise((resolve) => setTimeout(resolve, 140));
  });
  expect(onFind).toHaveBeenCalledTimes(1);
  await act(async () => root.render(wrap(<FindReplaceDialog {...props} isOpen={false} />)));
  expect(onClearHighlights).toHaveBeenCalled();
  await act(async () =>
    root.render(wrap(<FindReplaceDialog {...props} isOpen initialSearchText="reopened" />)),
  );
  expect(container.querySelector("input")?.value).toBe("reopened");
  await act(async () => root.unmount());
});

test("query edits invalidate unchanged host results and reopening an empty query has no matches", async () => {
  const container = document.createElement("div");
  const root = createRoot(container);
  const onFindNext = mock(() => externalResult.matches.at(0) ?? null);
  const onFindPrevious = mock(() => externalResult.matches.at(0) ?? null);
  const onFind = mock(() => externalResult);
  const props = {
    onClose: noop,
    onFind,
    onFindNext,
    onFindPrevious,
    onReplace: () => false,
    onReplaceAll: () => 0,
    currentResult: externalResult,
  };
  const render = (isOpen: boolean, initialSearchText: string) =>
    wrap(<FindReplaceDialog {...props} isOpen={isOpen} initialSearchText={initialSearchText} />);
  const navigation = () => [
    ...container.querySelectorAll<HTMLButtonElement>(
      'button[aria-label="Previous match"], button[aria-label="Next match"]',
    ),
  ];
  const assertReset = () => {
    expect(container.textContent).not.toMatch(/\d+ \/ \d+/u);
    expect(navigation()).toHaveLength(2);
    expect(navigation().every((button) => button.disabled)).toBe(true);
  };
  try {
    await act(async () => root.render(render(true, "selected")));
    assertReset();
    await act(async () => {
      await new Promise((resolve) => setTimeout(resolve, 140));
    });
    expect(container.textContent).toContain("2 / 3");
    const input = container.querySelector("input");
    if (!input) panic("Find input missing");
    // A new nonempty query must not navigate matches from the previous query
    // during the debounce interval, even when currentResult is unchanged.
    await act(async () => {
      input.value = "different";
      input.dispatchEvent(new Event("input", { bubbles: true }));
    });
    expect(input.value).toBe("different");
    assertReset();
    await act(async () =>
      root.render(
        wrap(
          <FindReplaceDialog
            {...props}
            isOpen
            initialSearchText="selected"
            currentResult={lateExternalResult}
          />,
        ),
      ),
    );
    // A new host identity for the old query cannot revive a cleared result.
    assertReset();
    await act(async () => {
      input.value = "";
      input.dispatchEvent(new Event("input", { bubbles: true }));
    });
    expect(input.value).toBe("");
    expect(onFind).toHaveBeenCalledWith("", { matchCase: false, matchWholeWord: false });
    assertReset();
    await act(async () => {
      for (const button of navigation()) button.click();
    });
    expect(onFindNext).not.toHaveBeenCalled();
    expect(onFindPrevious).not.toHaveBeenCalled();
    await act(async () => root.render(render(false, "")));
    await act(async () => root.render(render(true, "")));
    expect(container.querySelector("input")?.value).toBe("");
    assertReset();
    await act(async () => {
      for (const button of navigation()) button.click();
    });
    expect(onFindNext).not.toHaveBeenCalled();
    expect(onFindPrevious).not.toHaveBeenCalled();
  } finally {
    await act(async () => root.unmount());
  }
});

test("popup resets edit mode when the link changes and after closing", async () => {
  const container = document.createElement("div");
  document.body.append(container);
  const anchorEl = document.createElement("a");
  document.body.append(anchorEl);
  const root = createRoot(container);
  const props = { onNavigate: noop, onCopy: noop, onEdit: noop, onRemove: noop, onClose: noop };
  const first = popupData(anchorEl, "first");
  const second = popupData(anchorEl, "second");
  await act(async () => root.render(wrap(<HyperlinkPopup {...props} data={first} />)));
  const edit = container.querySelector<HTMLButtonElement>('button[title="Edit link"]');
  expect(edit).not.toBeNull();
  await act(async () => edit?.click());
  expect(container.querySelector("input")?.value).toBe("first");
  await act(async () => root.render(wrap(<HyperlinkPopup {...props} data={second} />)));
  expect(container.querySelector("input")).toBeNull();
  expect(container.textContent).toContain("https://second.example");
  await act(async () => root.render(wrap(<HyperlinkPopup {...props} data={null} />)));
  await act(async () => root.render(wrap(<HyperlinkPopup {...props} data={first} />)));
  expect(container.querySelector("input")).toBeNull();
  await act(async () => root.unmount());
  anchorEl.remove();
  container.remove();
});
