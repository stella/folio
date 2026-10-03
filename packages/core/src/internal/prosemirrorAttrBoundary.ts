import { panic } from "better-result";
import type { Node as PMNode } from "prosemirror-model";

export type ProseMirrorAttrIssue = {
  readonly path: string;
  readonly message: string;
};

export type ReadProseMirrorAttrsResult<T> =
  | { ok: true; value: T }
  | { ok: false; issues: readonly ProseMirrorAttrIssue[] };

export const attrsRecord = (attrs: unknown): Record<string, unknown> => {
  if (isRecord(attrs)) {
    return attrs;
  }

  return {};
};

export const isRecord = (value: unknown): value is Record<string, unknown> =>
  value !== null && typeof value === "object" && !Array.isArray(value);

export const attrsResult = <T>(
  attrs: Record<string, unknown>,
  issues: ProseMirrorAttrIssue[],
): ReadProseMirrorAttrsResult<T> => {
  if (issues.length > 0) {
    return { ok: false, issues };
  }

  let normalizedAttrs: T | undefined;
  return {
    ok: true,
    get value(): T {
      if (normalizedAttrs !== undefined) {
        return normalizedAttrs;
      }

      const presentAttrs: Record<string, unknown> = {};
      for (const key in attrs) {
        if (!Object.hasOwn(attrs, key)) {
          continue;
        }
        const value = attrs[key];
        if (value !== null) {
          presentAttrs[key] = value;
        }
      }

      // SAFETY: this module is the ProseMirror FFI boundary. The checks above
      // validate the attrs this code relies on before exposing the typed shape,
      // and null ProseMirror defaults are normalized to absent optional fields.
      normalizedAttrs = presentAttrs as T;
      return normalizedAttrs;
    },
    set value(value: T) {
      normalizedAttrs = value;
    },
  };
};

export const expectAttrs = <T>(result: ReadProseMirrorAttrsResult<T>, label: string): T => {
  if (result.ok) {
    return result.value;
  }

  const details = result.issues.map((issue) => `${issue.path}: ${issue.message}`).join("\n");
  panic(`Invalid ProseMirror ${label}:\n${details}`);
};

export function expectCachedNodeAttrs<T extends object>(
  node: PMNode,
  cache: WeakMap<PMNode, T>,
  reader: (node: PMNode) => ReadProseMirrorAttrsResult<T>,
  label: string,
): T {
  const cached = cache.get(node);
  if (cached) {
    return cached;
  }

  const value = expectAttrs(reader(node), label);
  cache.set(node, value);
  return value;
}

export const expectNodeType = (
  node: PMNode,
  expected: string,
  issues: ProseMirrorAttrIssue[],
): void => {
  if (node.type.name !== expected) {
    issues.push({
      path: "node.type.name",
      message: `Expected ${expected}, got ${node.type.name}.`,
    });
  }
};
