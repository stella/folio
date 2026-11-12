// Each read below uses a block id held across a resolution: 7 reports.
const settle = () => {
  reviewer.acceptAll();
};

export const linear = () => {
  const id = reviewer.snapshot().blocks.at(0)?.id;
  reviewer.acceptAll();
  apply(id);
};

export const viaIdOf = () => {
  const id = idOf(reviewer, "Anchor");
  reviewer.rejectAll();
  apply(id);
};

export const derived = () => {
  const ids = { anchor: idOf(reviewer, "Anchor") };
  reviewer.acceptChange(change);
  apply(ids.anchor);
};

export const inALoop = () => {
  const id = idOf(reviewer, "Anchor");
  for (const batch of batches) {
    apply(id, batch);
    reviewer.acceptAll();
  }
};

export const throughALocalFunction = () => {
  const id = idOf(reviewer, "Anchor");
  settle();
  apply(id);
};

export const captured = () => {
  const id = idOf(reviewer, "Anchor");
  const use = () => apply(id);
  reviewer.rejectChange(change);
  use();
};

export const heldThenResolved = () => {
  const held = heldBlockIds(reviewer, { anchor: "Anchor" });
  reviewer.acceptSuggestion(suggestionId);
  apply(held("anchor"));
};
