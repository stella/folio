import { isDeepStrictEqual } from "node:util";

type FirstDifferingOpPartOptions = {
  control: ReadonlyMap<string, Uint8Array>;
  edited: ReadonlyMap<string, Uint8Array>;
};

/** Stable archive ordering; content, missing parts and extra parts all count. */
export const firstDifferingOpPart = ({ control, edited }: FirstDifferingOpPartOptions) => {
  const paths = [...new Set([...control.keys(), ...edited.keys()])].toSorted();
  const part = paths.find((path) => !isDeepStrictEqual(control.get(path), edited.get(path)));
  return part;
};
