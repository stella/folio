import * as v from "valibot";

import type { CanonicalRoomRecord } from "./canonicalRoomAuthority";
import { toWireJson } from "./canonicalRoomWire";
import type { CanonicalTabJournal } from "./canonicalTabRoom";

/*
 * A tab room's durable journal in the origin's IndexedDB: one database for
 * every room, records keyed by room and position, the leadership term keyed
 * by room. Each call is one transaction, so an append is stored whole or not
 * at all.
 */

const DATABASE = "folio-canonical-rooms";
const DATABASE_VERSION = 1;
const RECORDS = "records";
const TERMS = "terms";

const ROW = v.object({ record: v.unknown() });
const TERM_ROW = v.object({ term: v.pipe(v.number(), v.safeInteger()) });

const requested = <T>(request: IDBRequest<T>): Promise<T> =>
  new Promise((resolve, reject) => {
    request.addEventListener("success", () => resolve(request.result));
    request.addEventListener("error", () => reject(request.error));
  });

const committed = (transaction: IDBTransaction): Promise<void> =>
  new Promise((resolve, reject) => {
    transaction.addEventListener("complete", () => resolve());
    transaction.addEventListener("error", () => reject(transaction.error));
    transaction.addEventListener("abort", () => reject(transaction.error));
  });

/**
 * The durable journal of the tab room `room` (the room's name and the
 * snapshot id, so two documents never share one), in `factory`'s databases.
 */
export const openCanonicalIndexedDbJournal = (
  room: string,
  factory: IDBFactory = indexedDB,
): CanonicalTabJournal => {
  let opened: Promise<IDBDatabase> | undefined;
  const database = (): Promise<IDBDatabase> => {
    opened ??= new Promise((resolve, reject) => {
      const request = factory.open(DATABASE, DATABASE_VERSION);
      request.addEventListener("upgradeneeded", () => {
        const created = request.result;
        created.createObjectStore(RECORDS, { keyPath: ["room", "index"] });
        created.createObjectStore(TERMS, { keyPath: "room" });
      });
      request.addEventListener("success", () => resolve(request.result));
      request.addEventListener("error", () => {
        opened = undefined;
        reject(request.error);
      });
    });
    return opened;
  };
  const rows = IDBKeyRange.bound([room, 0], [room, Number.MAX_SAFE_INTEGER]);

  return {
    load: async () => {
      const transaction = (await database()).transaction(
        [RECORDS, TERMS],
        "readonly",
      );
      const [stored, term] = await Promise.all([
        requested(transaction.objectStore(RECORDS).getAll(rows)),
        requested(transaction.objectStore(TERMS).get(room)),
      ]);
      const records: unknown[] = [];
      for (const row of stored) {
        const parsed = v.safeParse(ROW, row);
        // A row that is not one is kept as is: the restore refuses it.
        records.push(parsed.success ? parsed.output.record : row);
      }
      const termRow = v.safeParse(TERM_ROW, term);
      return { records, term: termRow.success ? termRow.output.term : 0 };
    },
    storeTerm: async (term) => {
      const transaction = (await database()).transaction(TERMS, "readwrite");
      transaction.objectStore(TERMS).put({ room, term });
      await committed(transaction);
    },
    append: async (records: readonly CanonicalRoomRecord[]) => {
      const transaction = (await database()).transaction(RECORDS, "readwrite");
      const store = transaction.objectStore(RECORDS);
      const add = async (): Promise<void> => {
        const count = await requested(store.count(rows));
        for (const [offset, record] of records.entries())
          store.add({
            room,
            index: count + offset,
            record: toWireJson(record),
          });
      };
      await Promise.all([committed(transaction), add()]);
    },
  };
};
