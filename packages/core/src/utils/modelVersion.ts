type ModelSnapshot = {
  value: Record<string, unknown>;
  keys: string[];
  fields: unknown[];
  arrayLength: number | undefined;
  readToken: object;
};

type ModelGraphSnapshot = {
  nodes: ModelSnapshot[];
  byValue: WeakMap<object, ModelSnapshot>;
  version: object;
};

const isRecord = (value: unknown): value is Record<string, unknown> =>
  value !== null && typeof value === "object";

/** Exact versions of mutable graphs; warm reads compare a flat list of fields. */
export const createModelVersionTracker = () => {
  const graphs = new WeakMap<object, ModelGraphSnapshot>();

  return (value: unknown) => {
    if (!isRecord(value)) return value;
    const previous = graphs.get(value);
    const byValue = previous?.byValue ?? new WeakMap<object, ModelSnapshot>();
    const readToken = {};

    const inspect = (cached: ModelSnapshot) => {
      if (cached.readToken === readToken) return false;
      const { value: record } = cached;
      const keys = Object.keys(record);
      const arrayLength = Array.isArray(record) ? record.length : undefined;
      let changed = cached.arrayLength !== arrayLength || cached.keys.length !== keys.length;
      let fields = changed ? [] : cached.fields;
      for (let index = 0; index < keys.length; index += 1) {
        // SAFETY: index is bounded by keys.length.
        const key = keys[index]!;
        const field = record[key];
        if (!changed && (cached.keys[index] !== key || !Object.is(cached.fields[index], field))) {
          fields = cached.fields.slice(0, index);
          changed = true;
        }
        if (changed) fields.push(field);
      }
      if (changed) Object.assign(cached, { keys, fields, arrayLength });
      cached.readToken = readToken;
      return changed;
    };

    // Parents precede children. Stop at the first changed edge, so a detached
    // subgraph's getters are never read just because it used to be reachable.
    if (previous) {
      let changed = false;
      for (const node of previous.nodes) {
        if (inspect(node)) {
          // A later getter may throw while rebuilding. Never leave a mutated
          // snapshot associated with the previously published version.
          graphs.delete(value);
          changed = true;
          break;
        }
      }
      if (!changed) return previous.version;
    }

    // Rebuild only the reachable-node list, reusing already inspected values.
    // Each shared record/getter is read once, including during invalidation.
    const nodes: ModelSnapshot[] = [];
    const pending = [value];
    const visited = new Set<object>(pending);
    for (let index = 0; index < pending.length; index += 1) {
      // SAFETY: index is bounded by pending.length.
      const record = pending[index]!;
      let node = byValue.get(record);
      if (node) {
        inspect(node);
      } else {
        const keys = Object.keys(record);
        node = {
          value: record,
          keys,
          fields: keys.map((key) => record[key]),
          arrayLength: Array.isArray(record) ? record.length : undefined,
          readToken,
        };
        byValue.set(record, node);
      }
      nodes.push(node);
      for (const field of node.fields) {
        if (!isRecord(field) || visited.has(field)) continue;
        visited.add(field);
        pending.push(field);
      }
    }
    const version = {};
    graphs.set(value, { nodes, byValue, version });
    return version;
  };
};
