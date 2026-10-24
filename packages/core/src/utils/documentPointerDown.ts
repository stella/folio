type PointerDownListener = (event: PointerEvent) => void;

type DocumentPointerDownRegistration = {
  listeners: Set<PointerDownListener>;
  detach: () => void;
};

const registrations = new WeakMap<Document, DocumentPointerDownRegistration>();

export const onDocumentPointerDown = (
  ownerDocument: Document,
  listener: PointerDownListener,
) => {
  let existing = registrations.get(ownerDocument);
  if (!existing) {
    const listeners = new Set<PointerDownListener>();
    const handle = (event: PointerEvent) => {
      for (const callback of Array.from(listeners)) {
        if (listeners.has(callback)) callback(event);
      }
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
  const callback = (event: PointerEvent) => listener(event);
  registration.listeners.add(callback);
  return () => {
    if (!registration.listeners.delete(callback)) return;
    if (registration.listeners.size > 0) return;
    registration.detach();
    registrations.delete(ownerDocument);
  };
};
