import { describe, expect, mock, test } from "bun:test";
import { effectScope, nextTick, ref } from "vue";

import { useFontLifecycle, type FontLifecycleDependencies } from "./useFontLifecycle";

function deferred<T>() {
  return Promise.withResolvers<T>();
}

function createFontFace(): FontFace {
  return new FontFace("Document Sans", new ArrayBuffer(1));
}

describe("useFontLifecycle", () => {
  test("remeasures after embedded fonts load and removes them on document cleanup", async () => {
    const originalFontFace = Reflect.get(globalThis, "FontFace");
    const hadFontFace = "FontFace" in globalThis;
    Object.defineProperty(globalThis, "FontFace", {
      value: class {},
      configurable: true,
      writable: true,
    });

    try {
      const face = createFontFace();
      const loadEmbedded = mock((_buffer: ArrayBuffer) => Promise.resolve([face]));
      const remove = mock((_faces: readonly FontFace[]) => {});
      const remeasure = mock(() => {});
      const ready = ref(false);
      const buffer = new ArrayBuffer(8);
      const document = { originalBuffer: buffer };
      const dependencies: FontLifecycleDependencies = {
        loadEmbedded,
        loadHost: () => Promise.resolve([]),
        remove,
      };
      const scope = effectScope();
      scope.run(() => {
        useFontLifecycle(
          {
            isReady: ready,
            getDocument: () => document,
            fonts: () => undefined,
            remeasure,
          },
          dependencies,
        );
      });

      ready.value = true;
      await nextTick();
      await Promise.resolve();
      expect(loadEmbedded).toHaveBeenCalledWith(buffer);
      expect(remeasure).toHaveBeenCalledTimes(1);

      ready.value = false;
      await nextTick();
      expect(remove).toHaveBeenCalledWith([face]);
      scope.stop();
    } finally {
      if (hadFontFace) {
        Object.defineProperty(globalThis, "FontFace", {
          value: originalFontFace,
          configurable: true,
          writable: true,
        });
      } else {
        Reflect.deleteProperty(globalThis, "FontFace");
      }
    }
  });

  test("removes a stale async result without remeasuring the replacement document", async () => {
    const originalFontFace = Reflect.get(globalThis, "FontFace");
    const hadFontFace = "FontFace" in globalThis;
    Object.defineProperty(globalThis, "FontFace", {
      value: class {},
      configurable: true,
      writable: true,
    });

    try {
      const face = createFontFace();
      const pending = deferred<FontFace[]>();
      const remove = mock((_faces: readonly FontFace[]) => {});
      const remeasure = mock(() => {});
      const ready = ref(false);
      const document = { originalBuffer: new ArrayBuffer(8) };
      const scope = effectScope();
      scope.run(() => {
        useFontLifecycle(
          {
            isReady: ready,
            getDocument: () => document,
            fonts: () => undefined,
            remeasure,
          },
          {
            loadEmbedded: () => pending.promise,
            loadHost: () => Promise.resolve([]),
            remove,
          },
        );
      });

      ready.value = true;
      await nextTick();
      ready.value = false;
      await nextTick();
      pending.resolve([face]);
      await Promise.resolve();

      expect(remove).toHaveBeenCalledWith([face]);
      expect(remeasure).not.toHaveBeenCalled();
      scope.stop();
    } finally {
      if (hadFontFace) {
        Object.defineProperty(globalThis, "FontFace", {
          value: originalFontFace,
          configurable: true,
          writable: true,
        });
      } else {
        Reflect.deleteProperty(globalThis, "FontFace");
      }
    }
  });

  test("remeasures when host faces are added and removed", async () => {
    const originalFontFace = Reflect.get(globalThis, "FontFace");
    const hadFontFace = "FontFace" in globalThis;
    Object.defineProperty(globalThis, "FontFace", {
      value: class {},
      configurable: true,
      writable: true,
    });

    try {
      const face = createFontFace();
      const fonts = ref([{ family: "Host Sans", src: "/host-sans.woff2" }]);
      const loadHost = mock(() => Promise.resolve([face]));
      const remove = mock((_faces: readonly FontFace[]) => {});
      const remeasure = mock(() => {});
      const scope = effectScope();
      scope.run(() => {
        useFontLifecycle(
          {
            isReady: ref(false),
            getDocument: () => null,
            fonts: () => fonts.value,
            remeasure,
          },
          {
            loadEmbedded: () => Promise.resolve([]),
            loadHost,
            remove,
          },
        );
      });

      await Promise.resolve();
      expect(loadHost).toHaveBeenCalledWith(fonts.value);
      expect(remeasure).toHaveBeenCalledTimes(1);

      fonts.value = [];
      await nextTick();
      expect(remove).toHaveBeenCalledWith([face]);
      expect(remeasure).toHaveBeenCalledTimes(2);
      scope.stop();
    } finally {
      if (hadFontFace) {
        Object.defineProperty(globalThis, "FontFace", {
          value: originalFontFace,
          configurable: true,
          writable: true,
        });
      } else {
        Reflect.deleteProperty(globalThis, "FontFace");
      }
    }
  });
});
