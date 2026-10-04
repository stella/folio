import { expect, setDefaultTimeout, test } from "bun:test";
import { GlobalRegistrator } from "@happy-dom/global-registrator";
import fc from "fast-check";
import { EditorState } from "prosemirror-state";
import { applyDocumentOps, type FormattingPatch } from "@stll/docx-core/ops";
import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";
import type { Document, Paragraph, SectionProperties } from "../types/document";
import { CANONICAL_GAP, type CanonicalGap } from "../types/canonicalCapabilities";
import { createCanonicalSectionPropertiesOperation } from "./canonicalOperations";
import { createCanonicalSession, publishCanonicalProjection } from "./canonicalSession";
import { createHiddenEditorManager, type HiddenEditorManagerDeps } from "./hiddenEditorManager";

setDefaultTimeout(propertyTestTimeout(30_000));

const PRESENCE = ["absent", "undefined", "present"] as const;
const PATCH_KINDS = ["omitted", "undefined", "removed", "numbered"] as const;
type Presence = (typeof PRESENCE)[number];
type PatchKind = (typeof PATCH_KINDS)[number];

type FixtureOptions = { endpoints: number; finalPresence: Presence; notePresence: Presence };
const sectionFixture = ({ endpoints, finalPresence, notePresence }: FixtureOptions): Document => {
  const content: Paragraph[] = Array.from({ length: endpoints + 1 }, (_, index) => ({
    type: "paragraph",
    paraId: (index + 1).toString(16).padStart(8, "0"),
    content: [{ type: "run", content: [{ type: "text", text: `Section ${index}` }] }],
    ...(index < endpoints
      ? { sectionProperties: { pageWidth: 10000 + index, footnotePr: { numStart: index + 1 } } }
      : {}),
  }));
  const source: Document = { package: { document: { content } } };
  if (finalPresence === "undefined")
    Reflect.set(source.package.document, "finalSectionProperties", undefined);
  if (finalPresence === "present") {
    const properties: SectionProperties = { pageWidth: 12240, marginLeft: 700 };
    if (notePresence === "undefined") {
      Reflect.set(properties, "footnotePr", undefined);
      Reflect.set(properties, "endnotePr", undefined);
    }
    if (notePresence === "present") {
      properties.footnotePr = { position: "pageBottom", numFmt: "lowerLetter", numStart: 3 };
      properties.endnotePr = { position: "docEnd", numRestart: "eachSect", numStart: 5 };
    }
    source.package.document.finalSectionProperties = properties;
  }
  return source;
};

const notePatch = (kind: PatchKind, marginLeft: number): FormattingPatch<SectionProperties> => {
  switch (kind) {
    case "omitted":
      return { marginLeft };
    case "undefined":
      return { marginLeft, footnotePr: undefined, endnotePr: undefined };
    case "removed":
      return { marginLeft, footnotePr: null, endnotePr: null };
    case "numbered":
      return {
        marginLeft,
        footnotePr: {
          position: "beneathText",
          numFmt: "decimal",
          numStart: 9,
          numRestart: "continuous",
        },
        endnotePr: { position: "sectEnd", numFmt: "upperRoman", numStart: 4 },
      };
  }
};

const assertNotePatch = ({
  before,
  after,
  kind,
}: {
  before: SectionProperties | undefined;
  after: SectionProperties | undefined;
  kind: PatchKind;
}) => {
  expect(after).toBeDefined();
  for (const key of ["footnotePr", "endnotePr"] as const) {
    switch (kind) {
      case "omitted":
        expect(Object.hasOwn(after ?? {}, key)).toBe(Object.hasOwn(before ?? {}, key));
        expect(after?.[key]).toStrictEqual(before?.[key]);
        break;
      case "undefined":
        expect(Object.hasOwn(after ?? {}, key)).toBe(true);
        expect(after?.[key]).toBeUndefined();
        break;
      case "removed":
        expect(Object.hasOwn(after ?? {}, key)).toBe(false);
        break;
      case "numbered":
        expect(after?.[key]).toStrictEqual(notePatch(kind, 0)[key]);
        break;
    }
  }
};

test("generated final-section patches preserve earlier endpoints and exact note-property undo/redo", () => {
  const observed = new Set<PatchKind>();
  assertProperty(
    fc.property(
      fc.integer({ min: 0, max: 3 }),
      fc.constantFrom(...PRESENCE),
      fc.constantFrom(...PRESENCE),
      fc.array(fc.constantFrom(...PATCH_KINDS), { minLength: 1, maxLength: 8 }),
      (endpoints, finalPresence, notePresence, patches) => {
        const session = createCanonicalSession(
          sectionFixture({ endpoints, finalPresence, notePresence }),
        ).unwrap();
        let state = EditorState.create({ doc: session.projection.doc });
        const original = session.document;
        const originalText = state.doc.textContent;
        for (const [index, kind] of patches.entries()) {
          const before = session.document;
          const op = createCanonicalSectionPropertiesOperation(
            before,
            notePatch(kind, 800 + index),
          );
          expect(op.sectionIndex).toBe(endpoints);
          const prepared = session.prepareOperations(state, [op]).unwrap();
          const direct = applyDocumentOps(before, [op]).unwrap();
          const directUndo = applyDocumentOps(direct.document, direct.inverse);
          // The canonical journal also exercises its separately retained inverse below.
          expect(directUndo.unwrap().document).toStrictEqual(before);
          state = publishCanonicalProjection({ state, commit: prepared, session }).unwrap().state;
          assertNotePatch({
            before: before.package.document.finalSectionProperties,
            after: session.document.package.document.finalSectionProperties,
            kind,
          });
          expect(session.document.package.document.content).toStrictEqual(
            original.package.document.content,
          );
          expect(session.version).toBe(index + 1);
          observed.add(kind);
        }
        const edited = session.document;
        for (let index = 0; index < patches.length; index += 1) {
          const commit = session.prepareUndo(state).unwrap();
          state = publishCanonicalProjection({ state, commit, session }).unwrap().state;
        }
        expect(session.document).toStrictEqual(original);
        for (let index = 0; index < patches.length; index += 1) {
          const commit = session.prepareRedo(state).unwrap();
          state = publishCanonicalProjection({ state, commit, session }).unwrap().state;
        }
        expect(session.document).toStrictEqual(edited);
        expect(session.projection.doc.textContent).toBe(originalText);
      },
    ),
    { seed: 20261012, numRuns: 40 },
  );
  expect(observed).toEqual(new Set(PATCH_KINDS));
});

const managerDeps = ({
  host,
  source,
  readOnly,
  sessionKind,
  onRefusal,
  onDestroy = () => undefined,
}: {
  host: HTMLElement;
  source: Document;
  readOnly: () => boolean;
  sessionKind: "canonical" | "default";
  onRefusal: (message: string, gap: CanonicalGap) => void;
  onDestroy?: () => void;
}) =>
  ({
    getHost: () => host,
    getDocument: () => source,
    getStyles: () => null,
    getExtensionManager: () => undefined,
    getExternalPlugins: () => [],
    getCollaboration: () => undefined,
    getCollaborationModules: () => null,
    getPrecomputedInitialState: () => null,
    getReadOnly: readOnly,
    getExperimentalSession: () => (sessionKind === "canonical" ? "canonical" : undefined),
    getDocumentIdentity: () => "section-fixture",
    getDocumentContext: () => source,
    onTransaction: () => undefined,
    onSelectionChange: () => undefined,
    onKeyDown: () => false,
    onReadOnlyEditAttempt: () => undefined,
    onEditorViewReady: () => undefined,
    onEditorViewDestroy: onDestroy,
    onRemoteSelectionsChange: () => undefined,
    onSessionRefusal: onRefusal,
  }) satisfies HiddenEditorManagerDeps;

test("manager section patches refuse read-only and composition without changing model or journal", () => {
  GlobalRegistrator.register();
  try {
    assertProperty(
      fc.property(
        fc.integer({ min: 0, max: 3 }),
        fc.constantFrom(...PATCH_KINDS),
        (endpoints, kind) => {
          const host = document.createElement("div");
          document.body.append(host);
          const source = sectionFixture({
            endpoints,
            finalPresence: "present",
            notePresence: "present",
          });
          let readOnly = false;
          const refused: CanonicalGap[] = [];
          const manager = createHiddenEditorManager(
            managerDeps({
              host,
              source,
              readOnly: () => readOnly,
              sessionKind: "canonical",
              onRefusal: (_message, gap) => refused.push(gap),
            }),
          );
          try {
            manager.ensureView();
            const initial = manager.api.getCanonicalDocument();
            expect(initial).not.toBeNull();
            readOnly = true;
            const readonlyResult = manager.api.applyCanonicalSectionProperties(
              notePatch(kind, 900),
            );
            expect(readonlyResult?.status).toBe("refused");
            if (readonlyResult?.status === "refused")
              expect(readonlyResult.gap).toBe(CANONICAL_GAP.sectionProperties);
            expect(manager.api.getCanonicalDocument()).toStrictEqual(initial);
            expect(manager.api.canUndo()).toBe(false);
            readOnly = false;
            manager.api.updateCanonicalInputLifecycle("beginComposition");
            expect(manager.api.applyCanonicalSectionProperties(notePatch(kind, 900))?.status).toBe(
              "refused",
            );
            expect(manager.api.getCanonicalDocument()).toStrictEqual(initial);
            expect(manager.api.canUndo()).toBe(false);
            manager.api.updateCanonicalInputLifecycle("endComposition");
            expect(manager.api.applyCanonicalSectionProperties(notePatch(kind, 900))).toEqual({
              status: "applied",
              version: 1,
            });
            const edited = manager.api.getCanonicalDocument();
            expect(manager.api.undo()).toBe(true);
            expect(manager.api.getCanonicalDocument()).toStrictEqual(initial);
            expect(manager.api.redo()).toBe(true);
            expect(manager.api.getCanonicalDocument()).toStrictEqual(edited);
            expect(refused).toEqual([
              CANONICAL_GAP.sectionProperties,
              CANONICAL_GAP.sectionProperties,
            ]);
          } finally {
            manager.destroyView();
            host.remove();
          }
        },
      ),
      { seed: 20261013, numRuns: 12 },
    );
    const host = document.createElement("div");
    document.body.append(host);
    const manager = createHiddenEditorManager(
      managerDeps({
        host,
        source: sectionFixture({ endpoints: 1, finalPresence: "present", notePresence: "absent" }),
        readOnly: () => false,
        sessionKind: "default",
        onRefusal: () => undefined,
      }),
    );
    try {
      expect(manager.api.applyCanonicalSectionProperties({ marginLeft: 900 })).toBeNull();
    } finally {
      manager.destroyView();
      host.remove();
    }
  } finally {
    GlobalRegistrator.unregister();
  }
});

test("section API refuses teardown reentry with one typed refusal and no model mutation", () => {
  GlobalRegistrator.register();
  const host = document.createElement("div");
  document.body.append(host);
  let invokeDuringDestroy: (() => void) | undefined;
  const refused: CanonicalGap[] = [];
  const manager = createHiddenEditorManager(
    managerDeps({
      host,
      source: sectionFixture({ endpoints: 2, finalPresence: "present", notePresence: "present" }),
      readOnly: () => false,
      sessionKind: "canonical",
      onRefusal: (_message, gap) => refused.push(gap),
      onDestroy: () => invokeDuringDestroy?.(),
    }),
  );
  try {
    manager.ensureView();
    const initial = manager.api.getCanonicalDocument();
    expect(initial).not.toBeNull();
    let attempts = 0;
    invokeDuringDestroy = () => {
      attempts += 1;
      const result = manager.api.applyCanonicalSectionProperties({ marginLeft: 1900 });
      expect(result?.status).toBe("refused");
      if (result?.status === "refused") expect(result.gap).toBe(CANONICAL_GAP.sectionProperties);
      expect(manager.api.getCanonicalDocument()).toStrictEqual(initial);
      expect(manager.api.canUndo()).toBe(false);
    };
    manager.destroyView();
    expect(attempts).toBe(1);
    expect(refused).toEqual([CANONICAL_GAP.sectionProperties]);
  } finally {
    manager.destroyView();
    host.remove();
    GlobalRegistrator.unregister();
  }
});
