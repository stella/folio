// Each id is read again after the resolution, or used before it: 0 reports.
export const readAgain = () => {
  let id = idOf(reviewer, "Anchor");
  reviewer.acceptAll();
  id = idOf(reviewer, "Anchor");
  apply(id);
};

export const usedBefore = () => {
  const id = idOf(reviewer, "Anchor");
  apply(id);
  reviewer.acceptAll();
};

export const resolvedByIt = () => {
  const id = idOf(reviewer, "Anchor");
  reviewer.acceptChange(id);
};

export const readAgainInALoop = () => {
  let held = heldBlockIds(reviewer, { anchor: "Anchor" });
  for (const batch of batches) {
    batch.map(() => apply(held("anchor")));
    reviewer.acceptAll();
    held = heldBlockIds(reviewer, { anchor: "Anchor" });
  }
};

export const noResolution = () => {
  const id = idOf(reviewer, "Anchor");
  reviewer.toBuffer();
  apply(id);
};
