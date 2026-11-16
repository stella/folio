import { GlobalRegistrator } from "@happy-dom/global-registrator";

GlobalRegistrator.register();

import { act, createRef } from "react";
import { createRoot } from "react-dom/client";

import { panic } from "better-result";
import { afterAll, expect, test } from "bun:test";
import { IntlProvider } from "use-intl";

import { createDocx } from "@stll/folio-core/docx/rezip";
import { getFolioMessages } from "@stll/folio-core/i18n/messages";
import type { Comment } from "@stll/folio-core/types/content";
import type { Document } from "@stll/folio-core/types/document";

import { DocxEditor } from "./DocxEditor";
import type { DocxEditorRef } from "./DocxEditor.props";

const previousActEnvironment = Reflect.get(
  globalThis,
  "IS_REACT_ACT_ENVIRONMENT",
);
Reflect.set(globalThis, "IS_REACT_ACT_ENVIRONMENT", true);
// Happy DOM lays nothing out; give the panels row room for the comment column.
const clientWidth = Object.getOwnPropertyDescriptor(
  HTMLElement.prototype,
  "clientWidth",
);
Object.defineProperty(HTMLElement.prototype, "clientWidth", {
  configurable: true,
  get: () => 4000,
});
afterAll(() => {
  if (clientWidth)
    Object.defineProperty(HTMLElement.prototype, "clientWidth", clientWidth);
  Reflect.set(globalThis, "IS_REACT_ACT_ENVIRONMENT", previousActEnvironment);
  GlobalRegistrator.unregister();
});

const COMMENT_ID = 9;

const commentedDocument = (): Document => ({
  package: {
    document: {
      content: [
        {
          type: "paragraph",
          content: [
            { type: "commentRangeStart", id: COMMENT_ID },
            { type: "run", content: [{ type: "text", text: "Reviewed" }] },
            { type: "commentRangeEnd", id: COMMENT_ID },
            {
              type: "run",
              content: [{ type: "text", text: " text" }],
            },
          ],
        },
      ],
      comments: [
        {
          id: COMMENT_ID,
          author: "Owner",
          content: [
            {
              type: "paragraph",
              content: [
                { type: "run", content: [{ type: "text", text: "Check." }] },
              ],
            },
          ],
        },
      ],
    },
  },
});

test.each([true, false])(
  "sidebar comment actions publish thread changes only when editable (readOnly %p)",
  async (readOnly) => {
    const container = document.createElement("div");
    document.body.append(container);
    const root = createRoot(container);
    const editor = createRef<DocxEditorRef>();
    const bytes = await createDocx(commentedDocument());
    const changes: Comment[][] = [];
    const documents: Document[] = [];
    const props = {
      documentBuffer: bytes,
      readOnly,
      showToolbar: false,
      onCommentsChange: (next: Comment[]) => changes.push(next),
      onChange: (next: Document) => documents.push(next),
    };
    try {
      await act(async () =>
        root.render(
          <IntlProvider
            locale="en"
            timeZone="UTC"
            messages={getFolioMessages("en")}
          >
            <DocxEditor
              experimentalSession="legacy"
              accessibleName="Document editor"
              ref={editor}
              {...props}
            />
          </IntlProvider>,
        ),
      );
      await act(async () => editor.current?.loadDocumentBuffer(bytes));
      await act(async () => editor.current?.ensureEditorView({ focus: false }));
      changes.length = 0;
      documents.length = 0;
      const findCard = () =>
        container.querySelector<HTMLElement>(
          `.docx-comment-card[data-comment-id="${String(COMMENT_ID)}"]`,
        );
      for (let attempt = 0; attempt < 20 && !findCard(); attempt += 1)
        await act(
          async () => new Promise((resolve) => setTimeout(resolve, 25)),
        );
      const card = findCard() ?? panic("Comment card did not render");
      await act(async () => card.click());
      const resolve =
        card.querySelector<HTMLButtonElement>('button[title="Resolve"]') ??
        panic("Resolve control did not render");
      await act(async () => resolve.click());
      if (readOnly) {
        expect(changes).toEqual([]);
        expect(documents).toEqual([]);
      } else {
        expect(changes.at(-1)?.at(0)?.done).toBe(true);
      }
    } finally {
      await act(async () => root.unmount());
      container.remove();
    }
  },
);
