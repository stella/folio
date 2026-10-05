import { GlobalRegistrator } from "@happy-dom/global-registrator";

let ownedLeases = 0;

/** Borrow an ambient DOM; release an owned registration after its last driver. */
export const acquireHarnessDom = () => {
  if (typeof document !== "undefined" && ownedLeases === 0) return () => {};
  if (ownedLeases === 0) GlobalRegistrator.register();
  ownedLeases++;
  let released = false;
  return () => {
    if (released) return;
    released = true;
    ownedLeases--;
    if (ownedLeases === 0) void GlobalRegistrator.unregister();
  };
};
