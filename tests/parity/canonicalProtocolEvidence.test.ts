import { expect, test } from "bun:test";
import fc from "fast-check";
import { assertProperty, propertyTestTimeout } from "../../test/property-testing";
import { canonicalProtocolEvidence } from "./canonicalProtocolEvidence";

test(
  "canonical protocol evidence retains original errors and correlates evaluations without payloads",
  () => {
    assertProperty(
      fc.property(fc.integer({ min: 1 }), fc.uuid(), fc.string(), (id, sessionId, payload) => {
        const request = canonicalProtocolEvidence(
          `pw:protocol SEND ► ${JSON.stringify({
            id,
            sessionId,
            method: "Runtime.callFunctionOn",
            params: {
              objectId: "1.1.1",
              awaitPromise: true,
              arguments: [
                { value: "(bytes) => globalThis.__folioCanonical.load(bytes)" },
                { value: { payload } },
                { value: payload },
              ],
            },
          })} +1ms`,
        );
        expect(request).toEqual({
          type: "evidence",
          message: {
            id,
            sessionId,
            method: "Runtime.callFunctionOn",
            contextId: undefined,
            objectId: "1.1.1",
            awaitPromise: true,
            operation: "canonical-load",
          },
        });
        const error = { code: -32000, message: payload };
        expect(
          canonicalProtocolEvidence(
            `pw:protocol ◀ RECV ${JSON.stringify({ id, sessionId, error })}`,
          ),
        ).toEqual({ type: "evidence", message: { id, sessionId, error } });
        expect(
          canonicalProtocolEvidence(
            `pw:protocol ◀ RECV ${JSON.stringify({ id, sessionId, result: { result: { value: payload } } })}`,
          ),
        ).toEqual({ type: "discard" });
      }),
      {
        numRuns: 50,
        id: "canonical protocol evidence retains original errors and correlates evaluations without payloads",
      },
    );
  },
  propertyTestTimeout(5_000),
);

test("canonical protocol evidence retains context destruction and ordinary test output", () => {
  const message = {
    method: "Runtime.executionContextDestroyed",
    params: { executionContextId: 7, executionContextUniqueId: "context-7" },
    sessionId: "session-7",
  };
  expect(canonicalProtocolEvidence(`pw:protocol ◀ RECV ${JSON.stringify(message)}`)).toEqual({
    type: "evidence",
    message: { id: undefined, ...message },
  });
  expect(canonicalProtocolEvidence("test failed: exact history differs")).toEqual({
    type: "output",
    text: "test failed: exact history differs",
  });
});
