import { afterAll, afterEach, beforeAll, describe, expect, test } from "bun:test";
import { GlobalRegistrator } from "@happy-dom/global-registrator";
import { DRAWING_RAW_XML_MODES, type DrawingRawXmlMode } from "@stll/docx-core/model";
import { EditorState } from "prosemirror-state";
import { EditorView } from "prosemirror-view";

import { expectImageAttrs } from "./attrs";
import { commitImageResize } from "./imageCommit";
import { schema } from "./schema";
import type { ImageAttrs } from "./schema/nodes";

const views: EditorView[] = [];

const GROUP_RAW_XML = "<w:drawing><wp:inline/></w:drawing>";

/** A document whose single paragraph holds one image at position 1. */
const mountImage = (attrs: ImageAttrs): EditorView => {
  const doc = schema.node("doc", null, [
    schema.node("paragraph", null, [schema.node("image", attrs)]),
  ]);
  const mount = document.createElement("div");
  document.body.append(mount);
  const view = new EditorView(mount, { state: EditorState.create({ doc }) });
  views.push(view);
  return view;
};

beforeAll(() => GlobalRegistrator.register());

afterEach(() => {
  for (const view of views.splice(0)) {
    const mount = view.dom.parentElement;
    view.destroy();
    mount?.remove();
  }
  document.body.replaceChildren();
});

afterAll(() => GlobalRegistrator.unregister());

describe("image commits on classified drawings", () => {
  test("resizes an ordinary editable image", () => {
    const view = mountImage({ src: "data:image/png;base64,AA==", width: 40, height: 40 });

    expect(commitImageResize(view, 1, 80, 80)).toBe(1);

    const node = view.state.doc.nodeAt(1);
    if (!node) {
      throw new Error("Expected the image node to survive the resize");
    }
    expect(expectImageAttrs(node).width).toBe(80);
  });

  const classifiedModes: DrawingRawXmlMode[] = [
    DRAWING_RAW_XML_MODES.PRESERVE_ONLY,
    DRAWING_RAW_XML_MODES.PREVIEW_ONLY,
  ];

  test.each(classifiedModes)("declines to resize a %s drawing", (mode) => {
    const view = mountImage({
      src: "data:image/svg+xml,%3Csvg%2F%3E",
      width: 40,
      height: 40,
      _docxRawXml: GROUP_RAW_XML,
      _docxRawXmlMode: mode,
    });

    expect(commitImageResize(view, 1, 80, 80)).toBeNull();

    const node = view.state.doc.nodeAt(1);
    if (!node) {
      throw new Error("Expected the image node to be left in place");
    }
    const attrs = expectImageAttrs(node);
    expect(attrs.width).toBe(40);
    expect(attrs._docxRawXml).toBe(GROUP_RAW_XML);
  });
});
