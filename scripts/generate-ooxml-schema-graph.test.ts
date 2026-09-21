import { describe, expect, test } from "bun:test";
import { readdirSync } from "node:fs";
import path from "node:path";

import { validateExactObjectKeys } from "./lib/exact-object";

import corpusSourceManifest from "../corpus/sources.json";
import evidenceRecord from "../specifications/evidence/records/scoped-local-elements.json";
import evidenceSchema from "../specifications/evidence/schema.json";
import generatedGraph from "../specifications/generated/docx-transitional-schema.gen.json";
import sourceManifest from "../specifications/sources.json";
import { buildOoxmlSchemaGraph, type SchemaBundleInput } from "./generate-ooxml-schema-graph";

const fixtureDirectory = path.join(import.meta.dir, "fixtures", "specifications");

const loadFixtureBundle = async (): Promise<SchemaBundleInput> => {
  const paths = ["schema-graph-main.xsd", "schema-graph-shared.xsd"];
  const documents = new Map<string, string>();
  for (const fixturePath of paths) {
    // oxlint-disable-next-line no-await-in-loop -- fixture order is deterministic
    documents.set(fixturePath, await Bun.file(path.join(fixtureDirectory, fixturePath)).text());
  }
  return {
    documents,
    entrypoints: ["schema-graph-main.xsd"],
    profile: "transitional",
    sourceId: "fixture",
    sourceSha256: "0".repeat(64),
  };
};

describe("OOXML schema graph generation", () => {
  test("preserves imports, inheritance, compositors, scoped elements, attributes, and enums", async () => {
    const graph = buildOoxmlSchemaGraph([await loadFixtureBundle()]);

    expect(graph.documents.map(({ path: documentPath }) => documentPath)).toEqual([
      "schema-graph-main.xsd",
      "schema-graph-shared.xsd",
    ]);
    expect(graph.documents.at(0)?.imports).toEqual([
      {
        kind: "import",
        namespace: "urn:folio:test:shared",
        schemaLocation: "schema-graph-shared.xsd",
        target: "schema-graph-shared.xsd",
      },
    ]);

    const state = graph.symbols.find(({ id }) => id === "simpleType:{urn:folio:test:main}State");
    expect(state?.enumValues).toEqual(["on", "off"]);
    expect(graph.inheritance).toContainEqual({
      base: "{urn:folio:test:main}Base",
      derived: "complexType:{urn:folio:test:main}Derived",
      method: "extension",
    });

    const slot = graph.children.find(
      ({ name, owner }) => name === "slot" && owner === "complexType:{urn:folio:test:main}Derived",
    );
    expect(slot).toMatchObject({
      maxOccurs: "unbounded",
      minOccurs: "0",
      namespace: "urn:folio:test:main",
    });
    expect(graph.children).toContainEqual(
      expect.objectContaining({
        name: "label",
        namespace: "urn:folio:test:main",
        owner: slot?.id,
        type: "{http://www.w3.org/2001/XMLSchema}string",
      }),
    );
    expect(graph.attributes).toContainEqual(
      expect.objectContaining({
        name: "mode",
        namespace: "",
        owner: slot?.id,
        type: "{urn:folio:test:main}State",
        use: "required",
      }),
    );
    expect(
      graph.children.some(
        ({ name, owner }) => name === "root" && owner === "element:{urn:folio:test:main}root",
      ),
    ).toBe(false);
  });

  test("is deterministic across source map insertion order", async () => {
    const bundle = await loadFixtureBundle();
    const reversed = new Map([...bundle.documents.entries()].toReversed());
    expect(buildOoxmlSchemaGraph([bundle])).toEqual(
      buildOoxmlSchemaGraph([{ ...bundle, documents: reversed }]),
    );
  });

  test("the committed graph points at the current source digests", () => {
    expect(generatedGraph.schemaVersion).toBe(1);
    expect(generatedGraph.profile).toBe("docx-transitional");
    expect(generatedGraph.documents).toHaveLength(16);
    expect(generatedGraph.symbols.length).toBeGreaterThan(1_000);
    for (const generatedSource of generatedGraph.sources) {
      const manifestSource = sourceManifest.sources.find(({ id }) => id === generatedSource.id);
      if (manifestSource?.sourceType !== "artifact" || typeof manifestSource.sha256 !== "string") {
        throw new TypeError(`Missing artifact source ${generatedSource.id}`);
      }
      expect(generatedSource.sha256).toBe(manifestSource.sha256);
    }
  });
});

describe("behavior evidence format", () => {
  const repoRoot = path.join(import.meta.dir, "..");
  const recordsDirectory = path.join(repoRoot, "specifications/evidence/records");
  const recordFiles = readdirSync(recordsDirectory)
    .filter((file) => file.endsWith(".json"))
    .toSorted();

  test("the directory holds records", () => {
    expect(recordFiles).toContain("scoped-local-elements.json");
    expect(recordFiles.length).toBeGreaterThan(1);
  });

  test.each(recordFiles)("%s is a valid evidence record", async (file) => {
    const record = (await Bun.file(path.join(recordsDirectory, file)).json()) as Record<
      string,
      unknown
    >;

    // Every property the schema declares is required, so exact keys is the
    // additionalProperties check and the required check at once.
    const keys = validateExactObjectKeys(record, evidenceSchema.required, `evidence ${file}`);
    expect(keys.isOk()).toBe(true);

    expect(record["schemaVersion"]).toBe(evidenceSchema.properties.schemaVersion.const);
    expect(record["id"]).toBe(file.replace(/\.json$/u, ""));
    expect(String(record["id"])).toMatch(new RegExp(evidenceSchema.properties.id.pattern, "u"));
    expect(evidenceSchema.properties.claimType.enum).toContain(record["claimType"]);
    expect(evidenceSchema.properties.confidence.enum).toContain(record["confidence"]);
    expect(String(record["claim"]).length).toBeGreaterThan(0);
    for (const profile of record["profiles"] as string[]) {
      expect(evidenceSchema.properties.profiles.items.enum).toContain(profile);
    }

    const sourceIds = new Set([
      ...sourceManifest.sources.map(({ id }) => id),
      ...corpusSourceManifest.sources.map(({ id }) => id),
    ]);
    const sources = record["sources"] as { sourceId: string; locator: string }[];
    expect(sources.length).toBeGreaterThan(0);
    for (const { sourceId, locator } of sources) {
      expect(sourceIds).toContain(sourceId);
      expect(locator.length).toBeGreaterThan(0);
    }

    const fixtures = record["fixtures"] as { path: string; assertion: string }[];
    expect(fixtures.length).toBeGreaterThan(0);
    for (const fixture of fixtures) {
      expect(fixture.assertion.length).toBeGreaterThan(0);
      // oxlint-disable-next-line no-await-in-loop -- one small file per fixture
      expect(await Bun.file(path.join(repoRoot, fixture.path)).exists()).toBe(true);
    }
  });

  test("the record the schema graph fixture backs is still present", () => {
    expect(evidenceRecord.id).toBe("scoped-local-elements");
  });
});
