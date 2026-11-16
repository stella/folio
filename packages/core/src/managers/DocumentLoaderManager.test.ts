import { describe, expect, test } from "bun:test";

import {
  normalizeForOps,
  packageParagraphIds,
  validateOpsDocument,
} from "@stll/docx-core/ops";

import {
  CanonicalDocxInputError,
  prepareCanonicalDocumentInput,
  prepareCanonicalDocxInput,
} from "../docx/canonicalSessionInput";
import { getParagraphText } from "../docx/paragraphParser";
import { parseDocx } from "../docx/parser";
import { createDocx } from "../docx/rezip";
import type { Document } from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";
import { DocumentLoaderManager } from "./DocumentLoaderManager";
import type {
  DocumentLoaderCallbacks,
  DocumentLoadState,
} from "./DocumentLoaderManager";

type Recorded = {
  events: string[];
  identities: string[];
  loadStates: DocumentLoadState[];
  errors: Error[];
  history: { state: Document | null };
};

const makeCallbacks = (): {
  callbacks: DocumentLoaderCallbacks;
  recorded: Recorded;
} => {
  const history: Recorded["history"] = { state: null };
  const recorded: Recorded = {
    events: [],
    identities: [],
    loadStates: [],
    errors: [],
    history,
  };
  const callbacks: DocumentLoaderCallbacks = {
    getExperimentalSession: () => "legacy",
    history: {
      get state() {
        return history.state;
      },
      reset: (document) => {
        history.state = document;
        recorded.events.push("history.reset");
      },
    },
    onError: (error) => {
      recorded.errors.push(error);
    },
    onCompatibilityChange: undefined,
    onReset: () => {
      recorded.events.push("onReset");
    },
    setDocumentLoadState: (state) => {
      recorded.loadStates.push(state);
      recorded.events.push(`load:${state.status}`);
    },
    setLoadedDocumentIdentity: (identity) => {
      recorded.identities.push(identity);
      recorded.events.push("identity");
    },
  };
  return { callbacks, recorded };
};

/** Settles once the load in flight has committed or failed. */
const waitForLoad = async (recorded: Recorded): Promise<void> => {
  for (let attempt = 0; attempt < 200; attempt += 1) {
    const last = recorded.loadStates.at(-1);
    if (last !== undefined && last.status !== "loading") return;
    await new Promise((resolve) => setTimeout(resolve, 5));
  }
  throw new Error("The load did not settle.");
};

describe("DocumentLoaderManager", () => {
  test.each([undefined, "provided-password"])(
    "canonical byte loads refuse encrypted containers before ZIP normalization (%s)",
    async (password) => {
      const { callbacks, recorded } = makeCallbacks();
      callbacks.getExperimentalSession = () => "canonical";
      const manager = new DocumentLoaderManager(callbacks);
      const compoundHeader = new Uint8Array([
        0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1,
      ]);
      await manager.loadBuffer(compoundHeader, { password });
      expect(recorded.history.state).toBeNull();
      expect(recorded.errors).toHaveLength(1);
      expect(recorded.errors.at(0)).toBeInstanceOf(CanonicalDocxInputError);
      expect(recorded.errors.at(0)?.message).toBe(
        "Password-protected documents are unavailable in the experimental canonical session.",
      );
      expect(recorded.loadStates.at(-1)).toEqual({
        status: "error",
        message: recorded.errors.at(0)?.message,
      });
    },
  );

  test("canonical ZIP normalization failures preserve the typed error and cause", async () => {
    const { callbacks, recorded } = makeCallbacks();
    callbacks.getExperimentalSession = () => "canonical";
    const manager = new DocumentLoaderManager(callbacks);
    await manager.loadBuffer(new Uint8Array([1, 2, 3, 4]));
    expect(recorded.history.state).toBeNull();
    expect(recorded.errors).toHaveLength(1);
    const error = recorded.errors.at(0);
    expect(error).toBeInstanceOf(CanonicalDocxInputError);
    if (!(error instanceof CanonicalDocxInputError)) return;
    expect(error.cause).toBeInstanceOf(Error);
    expect(error.message).toBe(
      error.cause instanceof Error ? error.cause.message : undefined,
    );
    expect(recorded.loadStates.at(-1)).toEqual({
      status: "error",
      message: error.message,
    });
  });

  test("canonical byte loads establish unique paragraph IDs before parsing", async () => {
    const { callbacks, recorded } = makeCallbacks();
    callbacks.getExperimentalSession = () => "canonical";
    const document = createEmptyDocument({ initialText: "Plain text" });
    const bytes = await createDocx(document);
    const manager = new DocumentLoaderManager(callbacks);
    await manager.loadBuffer(bytes);
    expect(recorded.errors).toEqual([]);
    const loaded = recorded.history.state;
    expect(loaded).not.toBeNull();
    if (!loaded) return;
    expect(validateOpsDocument(normalizeForOps(loaded)).isOk()).toBe(true);
    expect(loaded.package.document.content.at(0)?.type).toBe("paragraph");
  });

  test("a canonical parsed-document load opens a model without paragraph ids with the ids its package gets", async () => {
    const { callbacks, recorded } = makeCallbacks();
    callbacks.getExperimentalSession = () => "canonical";
    const model = createEmptyDocument({ initialText: "Plain text" });
    const paragraph = model.package.document.content.at(0);
    if (paragraph?.type !== "paragraph") throw new Error("No paragraph");
    delete paragraph.paraId;
    const manager = new DocumentLoaderManager(callbacks);
    manager.loadParsedDocument(model);
    expect(recorded.loadStates).toEqual([{ status: "loading" }]);
    await waitForLoad(recorded);
    expect(recorded.errors).toEqual([]);
    const loaded = recorded.history.state;
    if (!loaded) throw new Error("Nothing loaded");
    expect(validateOpsDocument(normalizeForOps(loaded)).isOk()).toBe(true);
    const direct = await parseDocx(
      (await prepareCanonicalDocxInput(await createDocx(model))).unwrap(),
      { detectVariables: false, preloadFonts: false },
    );
    expect(packageParagraphIds(loaded.package)).toEqual(
      packageParagraphIds(direct.package),
    );
    expect(recorded.identities).toHaveLength(1);
  });

  test("a canonical parsed-document load of a model with unique paragraph ids lands at once, as it is", () => {
    const { callbacks, recorded } = makeCallbacks();
    callbacks.getExperimentalSession = () => "canonical";
    const model = createEmptyDocument({ initialText: "Plain text" });
    new DocumentLoaderManager(callbacks).loadParsedDocument(model);
    expect(recorded.history.state).toBe(model);
  });

  test("a buffer load supersedes a canonical parsed-document load still establishing ids", async () => {
    const { callbacks, recorded } = makeCallbacks();
    callbacks.getExperimentalSession = () => "canonical";
    const model = createEmptyDocument({ initialText: "Model" });
    const paragraph = model.package.document.content.at(0);
    if (paragraph?.type !== "paragraph") throw new Error("No paragraph");
    delete paragraph.paraId;
    const manager = new DocumentLoaderManager(callbacks);
    manager.loadParsedDocument(model);
    await manager.loadBuffer(
      await createDocx(createEmptyDocument({ initialText: "Bytes" })),
    );
    // The superseded preparation has finished by the time this one has.
    await prepareCanonicalDocumentInput(model);
    expect(recorded.identities).toHaveLength(1);
    const loaded = recorded.history.state;
    const first = loaded?.package.document.content.at(0);
    expect(first?.type === "paragraph" && getParagraphText(first)).toBe(
      "Bytes",
    );
  });

  test("every parsed-document load lands with a fresh identity in the same commit as history", () => {
    const { callbacks, recorded } = makeCallbacks();
    const manager = new DocumentLoaderManager(callbacks);
    const first = createEmptyDocument();
    const second = createEmptyDocument();

    manager.loadParsedDocument(first);
    manager.loadParsedDocument(second);

    // Two loads of documents with identical metadata are still two distinct
    // external loads: the identity is per load, not per document signature.
    expect(recorded.identities).toHaveLength(2);
    expect(recorded.identities[0]).not.toBe(recorded.identities[1]);
    expect(recorded.history.state).toBe(second);
    // The identity is published right after the history reset and before the
    // load state flips to ready, so the adapter batches all three into one
    // render: hidden-editor resync and new document arrive together.
    expect(recorded.events).toEqual([
      "onReset",
      "history.reset",
      "identity",
      "load:ready",
      "onReset",
      "history.reset",
      "identity",
      "load:ready",
    ]);
  });

  test("a parsed-document load supersedes a buffer parse still in flight", async () => {
    const { callbacks, recorded } = makeCallbacks();
    const manager = new DocumentLoaderManager(callbacks);
    const parsed = createEmptyDocument();

    // Not a DOCX: the parse rejects, but only after the parsed load below has
    // landed. Without generation ordering across both entry points the stale
    // buffer outcome would clobber the newer document (here: an error state).
    const inFlight = manager.loadBuffer(new Uint8Array([1, 2, 3, 4]).buffer);
    manager.loadParsedDocument(parsed);
    await inFlight;

    expect(recorded.errors).toEqual([]);
    expect(recorded.history.state).toBe(parsed);
    expect(recorded.identities).toHaveLength(1);
    expect(recorded.loadStates.at(-1)).toEqual({ status: "ready" });
  });
});
