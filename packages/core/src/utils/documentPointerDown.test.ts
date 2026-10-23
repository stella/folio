import { GlobalRegistrator } from "@happy-dom/global-registrator";

if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { expect, mock, spyOn, test } from "bun:test";

import { onDocumentPointerDown } from "./documentPointerDown";

test.each(["mouse", "touch", "pen"])(
  "captures a canceled %s press before the editor stops propagation",
  (pointerType) => {
    const ownerDocument = document.implementation.createHTMLDocument();
    const target = ownerDocument.createElement("div");
    ownerDocument.body.append(target);
    target.addEventListener("pointerdown", (event) => {
      event.preventDefault();
      event.stopPropagation();
    });
    const close = mock(() => {});
    const dispose = onDocumentPointerDown(ownerDocument, close);
    const press = new PointerEvent("pointerdown", {
      bubbles: true,
      cancelable: true,
      pointerType,
    });
    expect(target.dispatchEvent(press)).toBe(false);
    expect(close).toHaveBeenCalledTimes(1);
    dispose();
  },
);

test("shares one listener and releases it after the last subscription", () => {
  const ownerDocument = document.implementation.createHTMLDocument();
  const add = spyOn(ownerDocument, "addEventListener");
  const remove = spyOn(ownerDocument, "removeEventListener");
  const close = mock(() => {});
  const first = onDocumentPointerDown(ownerDocument, close);
  const second = onDocumentPointerDown(ownerDocument, close);
  expect(add).toHaveBeenCalledTimes(1);
  ownerDocument.dispatchEvent(new PointerEvent("pointerdown"));
  expect(close).toHaveBeenCalledTimes(2);
  first();
  first();
  expect(remove).not.toHaveBeenCalled();
  second();
  expect(remove).toHaveBeenCalledTimes(1);
  const replacement = onDocumentPointerDown(ownerDocument, close);
  first();
  ownerDocument.dispatchEvent(new PointerEvent("pointerdown"));
  expect(close).toHaveBeenCalledTimes(3);
  replacement();
  add.mockRestore();
  remove.mockRestore();
});
