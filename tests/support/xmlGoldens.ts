/**
 * Golden comparisons for the workspace's XML readers.
 *
 * Each golden holds what a reader answered, before folio read XML with its own
 * reader, for every repository fixture (`fixtures`, digests per package) and
 * for every hostile or malformed input in `xmlEdgeInputs.ts` (`edges`). A
 * reader must still give the same answer for every fixture. For the edge
 * inputs it must refuse everything it refused before, give the same answer
 * for everything it accepted before, and may refuse an input it used to
 * accept only where the test names that input and says why.
 */
import { expect } from "bun:test";
import { readFileSync } from "node:fs";

import { ZipArchive } from "@stll/docx-core/zip";

import { type FixturePackage } from "./xmlFixtureParts";

export type Answer = { value: unknown } | { thrown: string };

export const attempt = (run: () => unknown): Answer => {
  try {
    return { value: run() };
  } catch (error) {
    return { thrown: error instanceof Error ? error.message : String(error) };
  }
};

export const attemptAsync = async (
  run: () => Promise<unknown>,
): Promise<Answer> => {
  try {
    return { value: await run() };
  } catch (error) {
    return { thrown: error instanceof Error ? error.message : String(error) };
  }
};

export type GoldenSite = {
  fixtures: Record<string, unknown>;
  edges: Record<string, unknown>;
};

export const readGoldens = <Sites extends string>(
  path: string,
): Record<Sites, GoldenSite> => JSON.parse(readFileSync(path, "utf8"));

/** The fixture package with its `word/document.xml` replaced by `xml`. */
export const withDocument = async (
  base: FixturePackage,
  xml: string,
): Promise<Uint8Array<ArrayBuffer>> => {
  const zip = await ZipArchive.read(base.bytes);
  zip.set("word/document.xml", xml);
  return await zip.generate();
};

export const SAMPLE_FIXTURE = "tests/visual/fixtures/sample.docx";

type ExpectEdgesOptions = {
  golden: GoldenSite;
  answers: Readonly<Record<string, unknown>>;
  /** Whether an answer is this reader's refusal. */
  refused: (answer: unknown) => boolean;
  /** Inputs accepted before and refused now, each with the reason. */
  tightened?: Readonly<Record<string, string>>;
  /** Inputs accepted before and now, read differently, each with the reason. */
  reread?: Readonly<Record<string, string>>;
};

export const expectEdgesMatch = ({
  golden,
  answers,
  refused,
  tightened = {},
  reread = {},
}: ExpectEdgesOptions): void => {
  expect(Object.keys(answers).toSorted()).toEqual(
    Object.keys(golden.edges).toSorted(),
  );
  for (const [name, before] of Object.entries(golden.edges)) {
    const after = answers[name];
    if (refused(before)) {
      expect(refused(after), `${name} is still refused`).toBe(true);
    } else if (Object.hasOwn(tightened, name)) {
      expect(refused(after), `${name} is now refused`).toBe(true);
    } else if (Object.hasOwn(reread, name)) {
      expect(refused(after), `${name} is still accepted`).toBe(false);
      expect(after, `${name} reads differently`).not.toEqual(before);
    } else {
      expect(after, name).toEqual(before);
    }
  }
  for (const name of [...Object.keys(tightened), ...Object.keys(reread)]) {
    expect(refused(golden.edges[name]), `${name} was accepted before`).toBe(
      false,
    );
  }
};

export const expectFixturesMatch = (
  golden: GoldenSite,
  answers: Readonly<Record<string, unknown>>,
  /** Fixtures answered differently now, each with the reason. */
  reread: Readonly<Record<string, string>> = {},
): void => {
  expect(Object.keys(answers)).toEqual(Object.keys(golden.fixtures));
  for (const [path, before] of Object.entries(golden.fixtures)) {
    if (Object.hasOwn(reread, path)) {
      expect(answers[path], path).not.toEqual(before);
    } else {
      expect(answers[path], path).toEqual(before);
    }
  }
};

const isThrown = (answer: unknown): boolean =>
  typeof answer === "object" && answer !== null && "thrown" in answer;

const valueOf = (answer: unknown): unknown =>
  typeof answer === "object" && answer !== null && "value" in answer
    ? answer.value
    : undefined;

/** A refusal is a throw, or a value the predicate calls a refusal. */
export const refusedWhen =
  (refusal: (value: unknown) => boolean) =>
  (answer: unknown): boolean =>
    isThrown(answer) || refusal(valueOf(answer));

export const NEVER = (): boolean => false;
