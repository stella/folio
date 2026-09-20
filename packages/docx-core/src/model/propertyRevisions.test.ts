import { describe, expect, test } from "bun:test";

import { PROPERTY_REVISION_KINDS } from "./content";

const MODEL_SOURCE = await Bun.file(`${import.meta.dir}/content.ts`).text();

/**
 * One `export type …Change` declaration of the content model: the
 * discriminators it carries and whether it holds a `PropertyChangeInfo`,
 * which is what makes a change element a property revision.
 *
 * The census is read from the source rather than from the type system because
 * the question is exactly the one the type system cannot ask: whether a
 * declaration nobody added to {@link PROPERTY_REVISION_KINDS} exists.
 */
type ChangeDeclaration = {
  name: string;
  body: string;
  discriminators: string[];
  carriesPropertyChangeInfo: boolean;
};

const changeDeclarations = (): ChangeDeclaration[] => {
  const declarations: ChangeDeclaration[] = [];
  const header = /^export type (\w*Change) =/gm;
  for (const match of MODEL_SOURCE.matchAll(header)) {
    const name = match[1];
    if (name === undefined) {
      continue;
    }
    const start = match.index + match[0].length;
    const next = MODEL_SOURCE.indexOf("\nexport ", start);
    const body = MODEL_SOURCE.slice(start, next === -1 ? undefined : next);
    declarations.push({
      name,
      body,
      discriminators: [...body.matchAll(/^\s*type: "(\w+)"/gm)].flatMap(([, kind]) =>
        kind === undefined ? [] : [kind],
      ),
      carriesPropertyChangeInfo: /^\s*info: PropertyChangeInfo;/m.test(body),
    });
  }
  return declarations;
};

/**
 * `*Change` declarations that are NOT property revisions, and why. Naming them
 * is what makes the census total: a new declaration is a property revision, the
 * union over them, or an entry here, never an omission.
 */
const STRUCTURAL_CHANGE_DECLARATIONS: Record<string, string> = {
  TrackedRunChange: "a union of run wrappers that add or remove content, not a property snapshot",
  ParagraphMarkChange: "a paragraph break that was added or removed, keyed by `kind`",
};

const UNION_DECLARATION = "PropertyChange";

describe("property revision census", () => {
  test("every property revision the model declares is a known kind", () => {
    const declared = changeDeclarations();
    expect(declared.length).toBeGreaterThan(0);

    const revisions = declared.filter(({ carriesPropertyChangeInfo }) => carriesPropertyChangeInfo);
    expect(new Set(revisions.flatMap(({ discriminators }) => discriminators))).toEqual(
      new Set(Object.keys(PROPERTY_REVISION_KINDS)),
    );
  });

  test("the `PropertyChange` union names every property revision", () => {
    const declared = changeDeclarations();
    const union = declared.find(({ name }) => name === UNION_DECLARATION);
    if (!union) {
      throw new Error(`the model no longer declares \`${UNION_DECLARATION}\``);
    }
    const members = new Set([...union.body.matchAll(/\|\s*(\w+);?/g)].map(([, member]) => member));
    expect(members).toEqual(
      new Set(
        declared
          .filter(({ carriesPropertyChangeInfo }) => carriesPropertyChangeInfo)
          .map(({ name }) => name),
      ),
    );
  });

  test("every other `*Change` declaration is a named structural revision", () => {
    const unclassified = changeDeclarations().filter(
      ({ name, carriesPropertyChangeInfo }) =>
        !carriesPropertyChangeInfo &&
        name !== UNION_DECLARATION &&
        STRUCTURAL_CHANGE_DECLARATIONS[name] === undefined,
    );
    expect(unclassified.map(({ name }) => name)).toEqual([]);
  });

  test("the census keys its own values", () => {
    for (const [kind, value] of Object.entries(PROPERTY_REVISION_KINDS)) {
      expect(value).toBe(kind);
    }
  });
});
