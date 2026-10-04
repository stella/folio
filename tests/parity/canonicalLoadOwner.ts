type CanonicalLoadOwnerOptions = {
  previousOwner: object | null | undefined;
  getOwner: () => object | null | undefined;
  waitFrame: () => Promise<void>;
};

/** Parsing finishes before adapters adopt a new canonical session. */
export const waitForCanonicalLoadOwner = async ({
  previousOwner,
  getOwner,
  waitFrame,
}: CanonicalLoadOwnerOptions): Promise<void> => {
  const deadline = Date.now() + 10_000;
  for (;;) {
    const owner = getOwner();
    if (owner !== null && owner !== undefined && owner !== previousOwner) return;
    if (Date.now() >= deadline)
      throw new TypeError("Canonical loading did not establish a new owner");
    await waitFrame();
  }
};
