/**
 * Where invisible markers (comment ranges, bookmarks, permission ranges,
 * proofing marks) share one editor position, that position stands for the
 * source gap typed text takes there (`defaultInsertionGap`). A caret the
 * editor places and a caret an edit returns then name the same gap, so
 * characters typed one after another stay together.
 */
import { expect, test } from "bun:test";

import {
  defaultInsertionGap,
  OP_STORIES,
  paragraphLogicalText,
} from "@stll/docx-core/ops";

import { DOCUMENT_SHAPES } from "../__tests__/documentShapes";
import { parseDocx } from "../docx/parser";
import { createCanonicalSession } from "./canonicalSession";

const SHAPES = ["comments", "invisible-markers"] as const;

for (const id of SHAPES)
  test(`${id}: every insertion gap has an editor position, and it reads back as itself`, async () => {
    const shape = DOCUMENT_SHAPES.find((entry) => entry.id === id);
    if (shape === undefined) throw new TypeError(`Missing shape ${id}`);
    const document = await parseDocx(await shape.build(), {
      preloadFonts: false,
    });
    const session = createCanonicalSession(document).unwrap();
    const projection = session.projectStory(OP_STORIES.MAIN).unwrap();
    for (const block of document.package.document.content) {
      if (block.type !== "paragraph" || block.paraId === undefined) continue;
      const length = paragraphLogicalText(block).length;
      for (let offset = 0; offset <= length; offset += 1) {
        const address = {
          story: OP_STORIES.MAIN,
          blockId: block.paraId,
          offset,
          zeroWidthBefore: defaultInsertionGap(block.content, offset)
            .zeroWidthBefore,
        };
        const read = projection
          .addressAt(projection.positionAt(address).unwrap())
          .unwrap();
        expect(read.offset).toBe(offset);
        expect(read.zeroWidthBefore ?? 0).toBe(address.zeroWidthBefore);
      }
    }
  });
