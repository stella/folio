import { describe, expect, test } from "bun:test";

import { initialMode, resolveEditorAuthor } from "./editor-settings";

describe("resolveEditorAuthor", () => {
  test("prefers folio.author, then git, then the OS account", () => {
    expect(
      resolveEditorAuthor({ setting: " Ada ", gitUserName: "Git Name", osUserName: "ada" }),
    ).toBe("Ada");
    expect(resolveEditorAuthor({ setting: "", gitUserName: "Git Name", osUserName: "ada" })).toBe(
      "Git Name",
    );
    expect(resolveEditorAuthor({ setting: undefined, gitUserName: " ", osUserName: "ada" })).toBe(
      "ada",
    );
    expect(
      resolveEditorAuthor({ setting: undefined, gitUserName: undefined, osUserName: undefined }),
    ).toBe("Folio user");
  });
});

describe("initialMode", () => {
  test("edits directly unless tracked changes are on; read-only off disk", () => {
    expect(initialMode("off", true)).toBe("editing");
    expect(initialMode(undefined, true)).toBe("editing");
    expect(initialMode("on", true)).toBe("suggesting");
    expect(initialMode("on", false)).toBe("viewing");
  });
});
