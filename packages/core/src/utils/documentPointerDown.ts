type PointerDownListener = (event: PointerEvent) => void;

type DocumentPointerDownRegistration = {
  listeners: Set<PointerDownListener>;
  detach: () => void;
};

const registrations = new WeakMap<Document, DocumentPointerDownRegistration>();

export const onDocumentPointerDown = (
  ownerDocument: Document,
  listener: PointerDownListener,
): (() => void) => {
  let existing = registrations.get(ownerDocument);
  if (!existing) {
    const listeners = new Set<PointerDownListener>();
    const handle = (event: PointerEvent): void => {
      for (const callback of Array.from(listeners)) callback(event);
    };
    ownerDocument.addEventListener("pointerdown", handle, true);
    existing = {
      listeners,
      detach: () =>
        ownerDocument.removeEventListener("pointerdown", handle, true),
    };
    registrations.set(ownerDocument, existing);
  }
  const registration = existing;
  const callback = (event: PointerEvent): void => listener(event);
  registration.listeners.add(callback);
  return () => {
    if (!registration.listeners.delete(callback)) return;
    if (registration.listeners.size > 0) return;
    registration.detach();
    registrations.delete(ownerDocument);
  };
};
