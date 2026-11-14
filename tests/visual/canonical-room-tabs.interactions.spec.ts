import { createDocx } from "../../packages/core/src/docx/rezip";
import { createEmptyDocument } from "../../packages/core/src/utils/createDocument";
import type {} from "../parity/canonicalBridge";
import {
  reactPort,
  select,
  vuePort,
} from "../parity/canonicalInteractionAssertions";
import type {} from "../parity/canonicalRoomBridge";
import {
  type BrowserContext,
  expect,
  type Page,
  test,
} from "../support/unclassifiedProbe";

/*
 * Two tabs of one browser share a canonical room through the playgrounds'
 * `?collaboration=1` demo: the Web Lock leader, the BroadcastChannel and the
 * room's real IndexedDB journal. The tabs share one browser context, so they
 * share the origin's locks, channels and databases, and nothing else.
 */

const adapters = [
  { name: "React", port: reactPort },
  { name: "Vue", port: vuePort },
] as const;

const JOIN_TIMEOUT = 30_000;

const roomBridge = (page: Page) =>
  page.evaluate(() => {
    const bridge = globalThis.__folioCanonicalRoom;
    return bridge === undefined
      ? null
      : {
          ...bridge.status(),
          editable: bridge.editable(),
          document: bridge.document(),
          comments: bridge.comments(),
        };
  });

/** Open the room in a new tab and wait until its editor takes local edits. */
const openTab = async (context: BrowserContext, url: string) => {
  const page = await context.newPage();
  await page.goto(url);
  await waitJoined(page);
  return page;
};

const waitJoined = async (page: Page) => {
  await expect
    .poll(async () => (await roomBridge(page))?.editable ?? false, {
      timeout: JOIN_TIMEOUT,
    })
    .toBe(true);
};

/** Every tab holds the same body and comments; returns that document. */
const expectConverged = async (pages: readonly Page[]) => {
  let converged: string | null = null;
  await expect
    .poll(
      async () => {
        const documents = await Promise.all(
          pages.map(async (page) => (await roomBridge(page))?.document),
        );
        const [first] = documents;
        converged = typeof first === "string" ? first : null;
        return (
          converged !== null &&
          documents.every((document) => document === converged)
        );
      },
      { timeout: JOIN_TIMEOUT },
    )
    .toBe(true);
  if (converged === null) throw new Error("The tabs hold no document.");
  return converged as string;
};

const paragraphStart = async (page: Page, index: number) => {
  const position = await page.evaluate(
    (paragraph) =>
      globalThis.__folioCanonicalRoom?.paragraphStart(paragraph) ?? null,
    index,
  );
  if (position === null) throw new Error(`No paragraph ${index}.`);
  return position;
};

/** Type `text` with the keyboard at the start of paragraph `index`. */
const typeAt = async (page: Page, index: number, text: string) => {
  await select(page, await paragraphStart(page, index));
  await page.keyboard.type(text);
};

const leaderOf = async (pages: readonly Page[]) => {
  for (const page of pages)
    if ((await roomBridge(page))?.leading === true) return page;
  return undefined;
};

for (const adapter of adapters) {
  test(`two tabs of a canonical room converge through its journal [${adapter.name}]`, async ({
    context,
  }, testInfo) => {
    test.setTimeout(180_000);
    const room = `room-${adapter.name.toLowerCase()}-${testInfo.testId}-${Date.now()}`;
    const url = `http://localhost:${adapter.port}/?collaboration=1#${room}`;

    // The first tab seeds the room and leads it; the second joins.
    const first = await openTab(context, url);
    const second = await openTab(context, url);
    expect((await roomBridge(first))?.leading).toBe(true);
    expect((await roomBridge(second))?.leading).toBe(false);
    await expectConverged([first, second]);

    // Each tab sees the other's selection.
    await expect
      .poll(() => first.locator("[data-peer]").count(), {
        timeout: JOIN_TIMEOUT,
      })
      .toBeGreaterThan(0);

    // Typing in both tabs, in separate paragraphs, at once.
    await Promise.all([typeAt(first, 0, "alpha "), typeAt(second, 1, "beta ")]);
    const typed = await expectConverged([first, second]);
    expect(typed).toContain("alpha ");
    expect(typed).toContain("beta ");

    // A tracked change: the second tab suggests.
    expect(
      await second.evaluate(
        () => globalThis.__folioCanonical?.setMode("suggesting") ?? false,
      ),
    ).toBe(true);
    await typeAt(second, 1, "suggested ");
    const suggested = await expectConverged([first, second]);
    expect(suggested).toContain("suggested ");
    expect(suggested).toContain('"type":"insertion"');

    // A comment in the first tab, a reply to it in the second.
    const start = await paragraphStart(first, 0);
    const commentId = await first.evaluate(
      ({ from, to }) =>
        globalThis.__folioCanonicalRoom?.comment(from, to, "A comment") ?? null,
      { from: start, to: start + "alpha".length },
    );
    expect(commentId).not.toBeNull();
    await expect
      .poll(async () => (await roomBridge(second))?.comments?.length ?? 0, {
        timeout: JOIN_TIMEOUT,
      })
      .toBe(1);
    const replyId = await second.evaluate(
      (parentId) =>
        globalThis.__folioCanonicalRoom?.reply(parentId, "A reply") ?? null,
      commentId ?? -1,
    );
    expect(replyId).not.toBeNull();
    await expectConverged([first, second]);
    for (const page of [first, second]) {
      const comments = (await roomBridge(page))?.comments ?? [];
      expect(comments).toHaveLength(2);
      expect(comments).toEqual(
        expect.arrayContaining([
          { id: commentId, parentId: null, author: "Room tester", done: false },
          {
            id: replyId,
            parentId: commentId,
            author: "Room tester",
            done: false,
          },
        ]),
      );
    }

    // A reloaded tab rejoins as the same actor and catches up.
    const beforeReload = await expectConverged([first, second]);
    await second.reload();
    await waitJoined(second);
    expect(await expectConverged([first, second])).toBe(beforeReload);
    await typeAt(second, 1, "again ");
    expect(await expectConverged([first, second])).toContain("again ");

    // The leader closes: the other tab takes over from the journal.
    const leader = await leaderOf([first, second]);
    if (leader === undefined) throw new Error("The room has no leader.");
    const survivor = leader === first ? second : first;
    await leader.close();
    await expect
      .poll(async () => (await roomBridge(survivor))?.leading ?? false, {
        timeout: JOIN_TIMEOUT,
      })
      .toBe(true);
    await waitJoined(survivor);
    await typeAt(survivor, 0, "after ");
    // A tab that joins now sees only what the room sequenced and stored.
    const witness = await openTab(context, url);
    const afterHandover = await expectConverged([survivor, witness]);
    expect(afterHandover).toContain("after ");

    // With every tab closed, a new tab rebuilds the room from IndexedDB.
    await survivor.close();
    await witness.close();
    const reopened = await openTab(context, url);
    expect((await roomBridge(reopened))?.leading).toBe(true);
    expect(await expectConverged([reopened])).toBe(afterHandover);
    expect((await roomBridge(reopened))?.notice).toBeUndefined();
    await reopened.close();
  });
}

/** A body paragraph with a footnote reference, and the footnote. */
const noteDocument = async () => {
  const document = createEmptyDocument({ initialText: "Body seed" });
  const body = document.package.document.content.at(0);
  if (body?.type !== "paragraph")
    throw new TypeError("Expected a body paragraph.");
  body.content.push({
    type: "run",
    content: [{ type: "footnoteRef", id: 1 }],
  });
  document.package.footnotes = [
    {
      type: "footnote",
      id: 1,
      content: [
        {
          type: "paragraph",
          content: [
            { type: "run", content: [{ type: "text", text: "Footnote seed" }] },
          ],
        },
      ],
    },
  ];
  return [...new Uint8Array(await createDocx(document))];
};

/** Open the footnote's editor in `page`; returns its ProseMirror. */
const openFootnote = async (page: Page) => {
  const reference = page
    .locator('[data-note-kind="footnote"][data-note-id="1"]')
    .first();
  await expect(reference).toBeVisible({ timeout: JOIN_TIMEOUT });
  await reference.dblclick();
  const editor = page
    .locator(
      'aside[aria-label="Footnotes"], aside[aria-label="Endnotes"], .docx-editor-vue__note-editor',
    )
    .locator('[data-note-kind="footnote"][data-note-id="1"] .ProseMirror');
  await expect(editor).toBeVisible();
  return editor;
};

for (const adapter of adapters) {
  test(`an open note editor follows a peer's edit to its note [${adapter.name}]`, async ({
    context,
  }, testInfo) => {
    test.setTimeout(120_000);
    await context.addInitScript(
      (bytes) => {
        globalThis.__folioCanonicalRoomSeed = bytes;
      },
      await noteDocument(),
    );
    const room = `notes-${adapter.name.toLowerCase()}-${testInfo.testId}-${Date.now()}`;
    const url = `http://localhost:${adapter.port}/?collaboration=1#${room}`;
    const first = await openTab(context, url);
    const second = await openTab(context, url);
    await expectConverged([first, second]);

    const firstNote = await openFootnote(first);
    const secondNote = await openFootnote(second);
    // The first tab's caret sits in its footnote while the peer types there.
    await firstNote.click();
    await first.keyboard.press("End");

    // The second tab types in the footnote the first tab has open.
    await secondNote.click();
    await second.keyboard.press("End");
    await second.keyboard.type(" remote");
    await expect(secondNote).toContainText("Footnote seed remote");
    await expect(firstNote).toContainText("Footnote seed remote", {
      timeout: JOIN_TIMEOUT,
    });

    // The first tab's editor took the peer's edit: typing in it goes on.
    await firstNote.click();
    await first.keyboard.press("End");
    await first.keyboard.type(" local");
    await expect(firstNote).toContainText("Footnote seed remote local");
    await expect(secondNote).toContainText("Footnote seed remote local", {
      timeout: JOIN_TIMEOUT,
    });
    await expectConverged([first, second]);
    expect((await roomBridge(first))?.notice).toBeUndefined();
    expect((await roomBridge(second))?.notice).toBeUndefined();
  });
}
