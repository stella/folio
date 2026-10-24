import { afterAll, beforeAll, expect, mock, spyOn, test } from "bun:test";

import { acquireHarnessDom } from "../../../../test/harnessDom";
import { onDocumentPointerDown } from "./documentPointerDown";

let releaseDom: ReturnType<typeof acquireHarnessDom>;
beforeAll(() => {
  releaseDom = acquireHarnessDom();
});
afterAll(() => releaseDom());

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

test("honors listener membership while dispatching a snapshot", () => {
  const ownerDocument = document.implementation.createHTMLDocument();
  const calls: string[] = [];
  let disposeSecond = () => {};
  let disposeAdded = () => {};
  let added = false;
  const disposeFirst = onDocumentPointerDown(ownerDocument, () => {
    calls.push("first");
    disposeSecond();
    if (added) return;
    added = true;
    disposeAdded = onDocumentPointerDown(ownerDocument, () => {
      calls.push("added");
    });
  });
  disposeSecond = onDocumentPointerDown(ownerDocument, () => {
    calls.push("second");
  });

  ownerDocument.dispatchEvent(new PointerEvent("pointerdown"));
  expect(calls).toEqual(["first"]);
  ownerDocument.dispatchEvent(new PointerEvent("pointerdown"));
  expect(calls).toEqual(["first", "first", "added"]);

  disposeFirst();
  disposeAdded();
});
