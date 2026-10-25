import { acquireHarnessDom } from "../harnessDom";

export const withHarness = (run: () => void) => {
  const releaseDom = acquireHarnessDom();
  try {
    run();
  } finally {
    releaseDom();
  }
};

export const unrelatedRegistry = { register: () => {}, unregister: () => {} };
unrelatedRegistry.register();
unrelatedRegistry.unregister();
