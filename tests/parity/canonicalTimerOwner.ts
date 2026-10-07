/** The first caller of the instrumented timeout owns it; later frames do not. */
export const canonicalTimerOwner = (stack: string) => {
  const owner = stack.split("\n").at(2)?.trim();
  if (!owner?.startsWith("at ")) throw new TypeError("Timer capture has no owner frame");
  return owner;
};

export const isCanonicalInputTimer = (stack: string) =>
  /\/controller\/canonical(?:Composition|Input)\.ts:\d+:\d+/.test(canonicalTimerOwner(stack));
