const RECENT_LIMIT = 8;

/** Best-effort derivation candidates, newest first. */
export const recentValues = <Value extends object>() => {
  // A WeakRef target stays alive until the end of the current job. Hold one
  // weak list so evicted values can still be collected during a long batch.
  let reference: WeakRef<Value[]> | undefined;
  return {
    values: (): readonly Value[] => reference?.deref() ?? [],
    remember: (value: Value): void => {
      let values = reference?.deref();
      if (values === undefined) {
        values = [value];
        reference = new WeakRef(values);
        return;
      }
      const index = values.indexOf(value);
      if (index === 0) return;
      if (index !== -1) values.splice(index, 1);
      if (values.length === RECENT_LIMIT) values.pop();
      values.unshift(value);
    },
  };
};
