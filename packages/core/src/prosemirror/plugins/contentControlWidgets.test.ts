/**
 * Smoke tests for the content-control widgets plugin.
 *
 * The plugin's primary side-effect is dispatching transactions when the
 * user clicks inside a typed control. We exercise the helper functions
 * directly here (the click handler runs through the same code path) and
 * defer end-to-end DOM tests to Playwright.
 */

import { describe, expect, test } from "bun:test";
import { EditorState } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";

import {
  ContentControlBoundError,
  ContentControlLockedError,
  ContentControlTypeError,
} from "../../content-controls";

/** A minimal `<w:sdtPr>` carrying a `<w:dataBinding>` — marks the control bound. */
const BOUND_RAW_PROPERTIES_XML = '<w:sdtPr><w:dataBinding w:xpath="/root/field"/></w:sdtPr>';
import { schema, singletonManager } from "../schema";
import {
  dispatchDatePick,
  dispatchDropdownPick,
  handleContentControlWidgetClick,
} from "./contentControlWidgets";
import type { ContentControlWidgetEvent } from "./contentControlWidgets";

type StubView = Pick<EditorView, "state" | "dispatch">;
type StateRef = { state: EditorState };

type ClickWidgetOptions = {
  checked?: boolean;
  listItems?: string;
  stateRef: StateRef;
  tag?: string;
  type: string;
};

function viewLike(stateRef: StateRef): StubView {
  return {
    state: stateRef.state,
    dispatch(tr) {
      stateRef.state = stateRef.state.apply(tr);
    },
  };
}

const clickWidget = ({
  checked,
  listItems,
  stateRef,
  tag = "control",
  type,
}: ClickWidgetOptions) => {
  const dataset = {
    ...(checked === undefined ? {} : { sdtChecked: String(checked) }),
    ...(listItems === undefined ? {} : { sdtListItems: listItems }),
    sdtPmPos: "0",
    sdtTag: tag,
    sdtType: type,
  };
  const anchor = {
    dataset,
    getBoundingClientRect: () => ({ bottom: 80, left: 24 }),
  };
  const target = {
    closest: (selector: string) => (selector === "[data-sdt-type]" ? anchor : null),
  };
  const events: ContentControlWidgetEvent[] = [];
  let prevented = false;
  const handled = handleContentControlWidgetClick({
    view: viewLike(stateRef),
    event: {
      target,
      preventDefault: () => {
        prevented = true;
      },
    },
    onEvent: (event) => events.push(event),
  });
  return { anchor, events, handled, prevented };
};

describe("handleContentControlWidgetClick", () => {
  test("routes painted descendant clicks through the typed dropdown event", () => {
    const listItems = JSON.stringify([
      { value: "ca", displayText: "California" },
      { value: "ny", displayText: "New York" },
    ]);
    const sdt = schema.node(
      "blockSdt",
      {
        sdtType: "dropdown",
        tag: "state",
        listItems,
      },
      [schema.node("paragraph", {}, [schema.text("California")])],
    );
    const stateRef = {
      state: EditorState.create({
        doc: schema.node("doc", null, [sdt]),
        schema,
        plugins: [...singletonManager.getPlugins()],
      }),
    };
    const { anchor, events, handled, prevented } = clickWidget({
      listItems,
      stateRef,
      tag: "state",
      type: "dropdown",
    });

    expect(handled).toBe(true);
    expect(prevented).toBe(true);
    expect(events).toEqual([
      {
        kind: "dropdownOpen",
        tag: "state",
        pmPos: 0,
        sdtType: "dropdown",
        anchor,
        listItemsJson: listItems,
      },
    ]);
  });

  test("toggles a painted checkbox through a document transaction", () => {
    const sdt = schema.node(
      "blockSdt",
      {
        checked: false,
        sdtType: "checkbox",
        tag: "approval",
      },
      [schema.node("paragraph", {}, [schema.text("☐")])],
    );
    const stateRef = {
      state: EditorState.create({
        doc: schema.node("doc", null, [sdt]),
        schema,
        plugins: [...singletonManager.getPlugins()],
      }),
    };

    const { events, handled, prevented } = clickWidget({
      checked: false,
      stateRef,
      tag: "approval",
      type: "checkbox",
    });

    expect(handled).toBe(true);
    expect(prevented).toBe(true);
    expect(events).toEqual([]);
    expect(stateRef.state.doc.firstChild?.attrs["checked"]).toBe(true);
    expect(stateRef.state.doc.firstChild?.firstChild?.textContent).toBe("☒");
  });

  for (const sdtType of ["dropdown", "date"] as const) {
    test(`refuses a locked ${sdtType} before opening its picker`, () => {
      const sdt = schema.node(
        "blockSdt",
        {
          lock: "contentLocked",
          sdtType,
          tag: "locked",
        },
        [schema.node("paragraph", {}, [])],
      );
      const stateRef = {
        state: EditorState.create({
          doc: schema.node("doc", null, [sdt]),
          schema,
          plugins: [...singletonManager.getPlugins()],
        }),
      };

      const { events, handled, prevented } = clickWidget({
        stateRef,
        tag: "locked",
        type: sdtType,
      });

      expect(handled).toBe(true);
      expect(prevented).toBe(true);
      expect(events).toHaveLength(1);
      expect(events.at(0)?.kind).toBe("refused");
      expect(events.at(0)?.error).toBeInstanceOf(ContentControlLockedError);
    });
  }

  test("surfaces a lock error thrown while toggling a checkbox", () => {
    const sdt = schema.node(
      "blockSdt",
      {
        checked: false,
        lock: "contentLocked",
        sdtType: "checkbox",
        tag: "locked",
      },
      [schema.node("paragraph", {}, [schema.text("☐")])],
    );
    const stateRef = {
      state: EditorState.create({
        doc: schema.node("doc", null, [sdt]),
        schema,
        plugins: [...singletonManager.getPlugins()],
      }),
    };

    const { events, handled } = clickWidget({
      checked: false,
      stateRef,
      tag: "locked",
      type: "checkbox",
    });

    expect(handled).toBe(true);
    expect(events).toHaveLength(1);
    expect(events.at(0)?.kind).toBe("refused");
    expect(events.at(0)?.error).toBeInstanceOf(ContentControlLockedError);
  });

  test("refuses instead of throwing when toggling a bound (data-bound) checkbox", () => {
    // An attacker DOCX can bind any typed control to XML data; the click
    // handler must surface ContentControlBoundError as a refusal, the same
    // as a lock or type mismatch, not let it escape as an uncaught error.
    const sdt = schema.node(
      "blockSdt",
      {
        checked: false,
        sdtType: "checkbox",
        tag: "bound",
        rawPropertiesXml: BOUND_RAW_PROPERTIES_XML,
      },
      [schema.node("paragraph", {}, [schema.text("☐")])],
    );
    const stateRef = {
      state: EditorState.create({
        doc: schema.node("doc", null, [sdt]),
        schema,
        plugins: [...singletonManager.getPlugins()],
      }),
    };

    const { events, handled } = clickWidget({
      checked: false,
      stateRef,
      tag: "bound",
      type: "checkbox",
    });

    expect(handled).toBe(true);
    expect(events).toHaveLength(1);
    expect(events.at(0)?.kind).toBe("refused");
    expect(events.at(0)?.error).toBeInstanceOf(ContentControlBoundError);
    // Nothing was applied — the control's state is untouched.
    expect(stateRef.state.doc.firstChild?.attrs["checked"]).toBe(false);
  });

  test("surfaces a type error when painted metadata disagrees with the document", () => {
    const sdt = schema.node(
      "blockSdt",
      {
        sdtType: "dropdown",
        tag: "state",
      },
      [schema.node("paragraph", {}, [schema.text("California")])],
    );
    const stateRef = {
      state: EditorState.create({
        doc: schema.node("doc", null, [sdt]),
        schema,
        plugins: [...singletonManager.getPlugins()],
      }),
    };

    const { events, handled } = clickWidget({
      checked: false,
      stateRef,
      tag: "state",
      type: "checkbox",
    });

    expect(handled).toBe(true);
    expect(events).toHaveLength(1);
    expect(events.at(0)?.kind).toBe("refused");
    expect(events.at(0)?.error).toBeInstanceOf(ContentControlTypeError);
  });
});

describe("dispatchDropdownPick", () => {
  test("writes the picked value as the displayText", () => {
    const sdt = schema.node(
      "blockSdt",
      {
        sdtType: "dropdown",
        tag: "state",
        listItems: JSON.stringify([
          { value: "ca", displayText: "California" },
          { value: "ny", displayText: "New York" },
        ]),
      },
      [schema.node("paragraph", {}, [schema.text("☐")])],
    );
    const initial = EditorState.create({
      doc: schema.node("doc", null, [sdt]),
      schema,
      plugins: [...singletonManager.getPlugins()],
    });
    const ref = { state: initial };
    // The helper expects an `EditorView`; only `state` and `dispatch` are
    // touched here, so a structural stub satisfies the contract. The PM
    // position of the SDT is 0 (it is the doc's first child).
    const ok = dispatchDropdownPick(viewLike(ref) as EditorView, 0, "ny");
    expect(ok).toBe(true);
    expect(ref.state.doc.firstChild?.firstChild?.textContent).toBe("New York");
  });
});

describe("dispatchDropdownPick — lock handling", () => {
  test("returns false instead of throwing when the picked control is locked", () => {
    // The click-time preflight should have caught this, but the doc could
    // have changed mid-picker. The dispatch helper must not let a
    // ContentControlLockedError escape into the React handler — that
    // would surface as an uncaught UI error and crash the picker shell.
    const sdt = schema.node(
      "blockSdt",
      {
        sdtType: "dropdown",
        tag: "state",
        lock: "contentLocked",
        listItems: JSON.stringify([{ value: "ca", displayText: "California" }]),
      },
      [schema.node("paragraph", {}, [schema.text("☐")])],
    );
    const initial = EditorState.create({
      doc: schema.node("doc", null, [sdt]),
      schema,
      plugins: [...singletonManager.getPlugins()],
    });
    const ref = { state: initial };
    expect(() => dispatchDropdownPick(viewLike(ref) as EditorView, 0, "ca")).not.toThrow();
    expect(dispatchDropdownPick(viewLike(ref) as EditorView, 0, "ca")).toBe(false);
  });
});

describe("dispatchDropdownPick — bound handling", () => {
  test("returns false instead of throwing when the picked control is bound", () => {
    const sdt = schema.node(
      "blockSdt",
      {
        sdtType: "dropdown",
        tag: "state",
        rawPropertiesXml: BOUND_RAW_PROPERTIES_XML,
        listItems: JSON.stringify([{ value: "ca", displayText: "California" }]),
      },
      [schema.node("paragraph", {}, [schema.text("☐")])],
    );
    const initial = EditorState.create({
      doc: schema.node("doc", null, [sdt]),
      schema,
      plugins: [...singletonManager.getPlugins()],
    });
    const ref = { state: initial };
    expect(() => dispatchDropdownPick(viewLike(ref) as EditorView, 0, "ca")).not.toThrow();
    expect(dispatchDropdownPick(viewLike(ref) as EditorView, 0, "ca")).toBe(false);
  });
});

describe("dispatchDatePick", () => {
  test("returns false instead of throwing when the picked control is locked", () => {
    const sdt = schema.node(
      "blockSdt",
      {
        sdtType: "date",
        tag: "effective",
        lock: "sdtContentLocked",
      },
      [schema.node("paragraph", {}, [])],
    );
    const initial = EditorState.create({
      doc: schema.node("doc", null, [sdt]),
      schema,
      plugins: [...singletonManager.getPlugins()],
    });
    const ref = { state: initial };
    expect(() => dispatchDatePick(viewLike(ref) as EditorView, 0, "2026-06-02")).not.toThrow();
    expect(dispatchDatePick(viewLike(ref) as EditorView, 0, "2026-06-02")).toBe(false);
  });

  test("returns false instead of throwing when the picked control is bound", () => {
    const sdt = schema.node(
      "blockSdt",
      {
        sdtType: "date",
        tag: "effective",
        rawPropertiesXml: BOUND_RAW_PROPERTIES_XML,
      },
      [schema.node("paragraph", {}, [])],
    );
    const initial = EditorState.create({
      doc: schema.node("doc", null, [sdt]),
      schema,
      plugins: [...singletonManager.getPlugins()],
    });
    const ref = { state: initial };
    expect(() => dispatchDatePick(viewLike(ref) as EditorView, 0, "2026-06-02")).not.toThrow();
    expect(dispatchDatePick(viewLike(ref) as EditorView, 0, "2026-06-02")).toBe(false);
  });

  test("writes the picked date into the control body", () => {
    const sdt = schema.node("blockSdt", { sdtType: "date", tag: "effective" }, [
      schema.node("paragraph", {}, []),
    ]);
    const initial = EditorState.create({
      doc: schema.node("doc", null, [sdt]),
      schema,
      plugins: [...singletonManager.getPlugins()],
    });
    const ref = { state: initial };
    const ok = dispatchDatePick(viewLike(ref) as EditorView, 0, "2026-06-02");
    expect(ok).toBe(true);
    expect(ref.state.doc.firstChild?.firstChild?.textContent).toBe("2026-06-02");
  });
});
