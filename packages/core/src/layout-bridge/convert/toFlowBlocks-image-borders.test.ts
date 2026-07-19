import { describe, expect, test } from "bun:test";

import type { ImageRun } from "../../layout-engine/types";
import { schema } from "../../prosemirror/schema";
import { toFlowBlocks } from "./toFlowBlocks";

describe("image border conversion", () => {
  test("carries image border attrs into inline layout runs", () => {
    const image = schema.nodes.image.create({
      src: "data:image/png;base64,",
      width: 100,
      height: 80,
      borderWidth: 2,
      borderColor: "currentColor",
      borderStyle: "dashed",
    });
    const doc = schema.node("doc", null, [
      schema.node("paragraph", null, [schema.text("Before "), image, schema.text(" after")]),
    ]);

    const paragraph = toFlowBlocks(doc).find((block) => block.kind === "paragraph");
    expect(paragraph?.kind).toBe("paragraph");
    if (!paragraph || paragraph.kind !== "paragraph") {
      return;
    }

    const imageRun = paragraph.runs.find((run): run is ImageRun => run.kind === "image");
    expect(imageRun).toMatchObject({
      borderWidth: 2,
      borderColor: "currentColor",
      borderStyle: "dashed",
    });
  });
});
