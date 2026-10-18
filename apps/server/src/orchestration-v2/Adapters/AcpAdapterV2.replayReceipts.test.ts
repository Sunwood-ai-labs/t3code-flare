import { assert, describe, it } from "@effect/vitest";
import { ProviderDriverKind, type ProviderReplayEntry } from "@t3tools/contracts";
import * as Effect from "effect/Effect";
import * as Deferred from "effect/Deferred";
import * as Fiber from "effect/Fiber";
import * as Exit from "effect/Exit";
import * as AcpErrors from "effect-acp/errors";
import type * as AcpSchema from "effect-acp/compat";
import { ACP_PROTOCOL } from "@t3tools/provider-acp/server/adapter";
import { makeProviderReplayGate } from "@t3tools/provider-testing/replayGate";

import {
  type AcpReplayTranscript,
  withReplaySessionUpdateReceipts,
} from "./AcpAdapterV2.testkit.ts";

const notification = (
  sessionId = "session",
  eventId = "last-event",
): AcpSchema.SessionNotification => ({
  sessionId,
  update: { sessionUpdate: "agent_message_chunk", content: { type: "text", text: "3" } },
  _meta: { eventId },
});
const entry = (
  label = "last-output",
  sessionId = "session",
  eventId = "last-event",
): ProviderReplayEntry => ({
  type: "emit_inbound",
  label,
  frame: {
    kind: "notification",
    method: "session/update",
    params: notification(sessionId, eventId),
  },
});
const transcript = (
  entries: ReadonlyArray<ProviderReplayEntry> = [entry()],
): AcpReplayTranscript => ({
  provider: ProviderDriverKind.make("grok"),
  protocol: ACP_PROTOCOL,
  version: "1",
  scenario: "receipt",
  entries,
});

describe("ACP replay native-handler receipts", () => {
  it.effect("waits for the real final handler to return before acknowledging its event", () =>
    Effect.gen(function* () {
      const gate = makeProviderReplayGate([], ["last-output"]);
      const entered = yield* Deferred.make<void>();
      const pending = yield* Deferred.make<void>();
      let acknowledged = false;
      const receipt = gate.waitForHandled("last-output").then((value) => {
        acknowledged = value;
      });
      const handler = withReplaySessionUpdateReceipts(transcript(), gate, () =>
        Deferred.succeed(entered, undefined).pipe(Effect.andThen(Deferred.await(pending))),
      );
      const callback = yield* Effect.forkChild(handler(notification()));
      yield* Deferred.await(entered);
      yield* Effect.promise(() => Promise.resolve());
      assert.isFalse(acknowledged);
      yield* Deferred.succeed(pending, undefined);
      yield* Fiber.join(callback);
      yield* Effect.promise(() => receipt);
      assert.isTrue(acknowledged);
    }),
  );

  it.effect("does not acknowledge a failed real handler", () =>
    Effect.gen(function* () {
      const gate = makeProviderReplayGate([], ["last-output"]);
      let acknowledged = false;
      const receipt = gate.waitForHandled("last-output").then((value) => {
        acknowledged = value;
      });
      const handler = withReplaySessionUpdateReceipts(transcript(), gate, () =>
        Effect.fail(
          new AcpErrors.AcpTransportError({
            detail: "controlled callback failure",
            cause: "controlled",
          }),
        ),
      );
      const result = yield* Effect.exit(handler(notification()));
      assert.isTrue(Exit.isFailure(result));
      yield* Effect.promise(() => Promise.resolve());
      assert.isFalse(acknowledged);
      gate.recordHandled("last-output");
      yield* Effect.promise(() => receipt);
    }),
  );

  it.effect(
    "requires the same session and event, and an old event cannot satisfy another phase",
    () =>
      Effect.gen(function* () {
        const gate = makeProviderReplayGate([], ["last-output", "next-output"]);
        let lastAcknowledged = false;
        const last = gate.waitForHandled("last-output").then((value) => {
          lastAcknowledged = value;
        });
        let nextAcknowledged = false;
        const next = gate.waitForHandled("next-output").then((value) => {
          nextAcknowledged = value;
        });
        const handler = withReplaySessionUpdateReceipts(
          transcript([entry(), entry("next-output", "session", "next-event")]),
          gate,
          () => Effect.void,
        );
        yield* handler(notification("other-session"));
        yield* handler(notification("session", "other-event"));
        yield* Effect.promise(() => Promise.resolve());
        assert.isFalse(lastAcknowledged);
        yield* handler(notification());
        yield* Effect.promise(() => last);
        assert.isTrue(lastAcknowledged);
        yield* handler(notification());
        yield* Effect.promise(() => Promise.resolve());
        assert.isFalse(nextAcknowledged);
        yield* handler(notification("session", "next-event"));
        yield* Effect.promise(() => next);
        assert.isTrue(nextAcknowledged);
      }),
  );

  it.effect("allows the same event ID in another session without acknowledging the target", () =>
    Effect.gen(function* () {
      const gate = makeProviderReplayGate([], ["last-output"]);
      let acknowledged = false;
      const receipt = gate.waitForHandled("last-output").then((value) => {
        acknowledged = value;
      });
      const handler = withReplaySessionUpdateReceipts(
        transcript([entry("earlier-unwatched", "other-session"), entry()]),
        gate,
        () => Effect.void,
      );
      yield* handler(notification("other-session"));
      yield* Effect.promise(() => Promise.resolve());
      assert.isFalse(acknowledged);
      yield* handler(notification());
      yield* Effect.promise(() => receipt);
      assert.isTrue(acknowledged);
    }),
  );

  it("rejects ambiguous, missing or identity-free receipt targets", () => {
    const gate = makeProviderReplayGate([], ["last-output"]);
    for (const invalid of [
      transcript([]),
      transcript([entry(), entry()]),
      transcript([entry("earlier-unwatched"), entry()]),
      transcript([entry("last-output", "session", "")]),
    ]) {
      assert.throws(() => withReplaySessionUpdateReceipts(invalid, gate, () => Effect.void));
    }
    assert.throws(() =>
      withReplaySessionUpdateReceipts(
        transcript([entry(), entry("next-output")]),
        makeProviderReplayGate([], ["last-output", "next-output"]),
        () => Effect.void,
      ),
    );
  });
});
