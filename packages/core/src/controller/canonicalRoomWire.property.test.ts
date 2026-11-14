import { panic, Result } from "better-result";
import { expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import {
  BATCH_REJECTION_REASONS,
  BatchRejection,
  DOCUMENT_OP_TYPES,
  EXCLUSIVE_PARTS,
  ID_LEASE_KINDS,
  ID_LEASE_REFUSAL_REASONS,
  IdLeaseRefusal,
} from "@stll/docx-core/ops";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import {
  decodeLeaseGrant,
  decodeLeaseRenewal,
  decodeRoomSubmission,
  encodeLeaseGrant,
  encodeLeaseRenewal,
  encodeRoomSubmission,
  toWireJson,
  UNAVAILABLE_LEASE_ANSWER,
} from "./canonicalRoomWire";

setDefaultTimeout(propertyTestTimeout(30_000));

/*
 * What crosses between two contexts as JSON comes back as it left: every
 * field a rejection or a lease refusal carries survives the trip, and the
 * rebuilt values are the tagged errors again.
 */

const OP_TYPES = Object.values(DOCUMENT_OP_TYPES);
const LEASE_KINDS = Object.values(ID_LEASE_KINDS);

const rejectionFields = fc.record(
  {
    reason: fc.constantFrom(...Object.values(BATCH_REJECTION_REASONS)),
    message: fc.string(),
    opId: fc.string(),
    opType: fc.constantFrom(...OP_TYPES),
    overType: fc.constantFrom(...OP_TYPES),
    idKind: fc.constantFrom(...LEASE_KINDS),
    part: fc.constantFrom(...Object.values(EXCLUSIVE_PARTS)),
  },
  { requiredKeys: ["reason", "message"] },
);

const refusalFields = fc.record(
  {
    reason: fc.constantFrom(...Object.values(ID_LEASE_REFUSAL_REASONS)),
    message: fc.string(),
    kind: fc.constantFrom(...LEASE_KINDS),
  },
  { requiredKeys: ["reason", "message"] },
);

/** A tagged error's fields that are set, without its class. */
const fieldsOf = (error: object, keys: readonly string[]) =>
  Object.fromEntries(
    keys.flatMap((key) => {
      const value: unknown = Reflect.get(error, key);
      return value === undefined ? [] : [[key, value]];
    }),
  );
const REJECTION_KEYS = [
  "reason",
  "message",
  "opId",
  "opType",
  "overType",
  "idKind",
  "part",
] as const;
const REFUSAL_KEYS = ["reason", "message", "kind"] as const;

test("a rejection crosses the wire with every field it carries", async () => {
  await assertProperty(
    fc.property(
      rejectionFields,
      fc.string(),
      fc.nat(),
      (fields, opId, headRev) => {
        const decoded = decodeRoomSubmission(
          toWireJson(
            encodeRoomSubmission({
              type: "reject",
              opId,
              headRev,
              reason: new BatchRejection(fields),
            }),
          ),
        );
        if (decoded?.type !== "reject")
          return panic("The rejection did not come back as one.");
        expect(decoded.opId).toBe(opId);
        expect(decoded.headRev).toBe(headRev);
        expect(decoded.reason).toBeInstanceOf(BatchRejection);
        expect(fieldsOf(decoded.reason, REJECTION_KEYS)).toEqual(fields);
      },
    ),
    { id: "canonical room wire rejection", numRuns: 200 },
  );
});

test("a lease refusal crosses the wire with every field it carries", async () => {
  await assertProperty(
    fc.property(refusalFields, (fields) => {
      const refusal = new IdLeaseRefusal(fields);
      for (const decoded of [
        decodeLeaseGrant(toWireJson(encodeLeaseGrant(Result.err(refusal)))),
        decodeLeaseRenewal(toWireJson(encodeLeaseRenewal(Result.err(refusal)))),
      ]) {
        if (decoded === undefined || decoded.isOk())
          return panic("The refusal did not come back as one.");
        expect(decoded.error).toBeInstanceOf(IdLeaseRefusal);
        expect(fieldsOf(decoded.error, REFUSAL_KEYS)).toEqual(fields);
      }
    }),
    { id: "canonical room wire lease refusal", numRuns: 100 },
  );
});

test("an abandoned lease call reads as an unavailable refusal to both decoders", () => {
  for (const decoded of [
    decodeLeaseGrant(UNAVAILABLE_LEASE_ANSWER),
    decodeLeaseRenewal(UNAVAILABLE_LEASE_ANSWER),
  ])
    expect(decoded?.isErr() === true && decoded.error.reason).toBe(
      ID_LEASE_REFUSAL_REASONS.UNAVAILABLE,
    );
});
