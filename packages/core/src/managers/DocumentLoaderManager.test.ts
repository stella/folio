import { describe, expect, test } from "bun:test";

import type { Document } from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";
import { DocumentLoaderManager } from "./DocumentLoaderManager";
import type { DocumentLoaderCallbacks, DocumentLoadState } from "./DocumentLoaderManager";

type Recorded = {
  events: string[];
  identities: string[];
  loadStates: DocumentLoadState[];
  errors: Error[];
  history: { state: Document | null };
};

const makeCallbacks = (): { callbacks: DocumentLoaderCallbacks; recorded: Recorded } => {
  const history: Recorded["history"] = { state: null };
  const recorded: Recorded = { events: [], identities: [], loadStates: [], errors: [], history };
  const callbacks: DocumentLoaderCallbacks = {
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

describe("DocumentLoaderManager", () => {
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
