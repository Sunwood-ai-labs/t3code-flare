import {
  MessageId,
  type ModelSelection,
  NodeId,
  ProjectId,
  type ProviderInstanceId,
  RunAttemptId,
  RunId,
  type OrchestrationV2ProviderThread,
  type ThreadId,
  ProviderDriverKind,
  ProviderReplayEntry,
  type ProviderReplayTranscript,
} from "@t3tools/contracts";
import type * as DateTime from "effect/DateTime";
import type * as ProviderAdapter from "@t3tools/provider-core/server/ProviderAdapter";
import * as Effect from "effect/Effect";
import * as Crypto from "effect/Crypto";
import * as FileSystem from "effect/FileSystem";
import * as Layer from "effect/Layer";
import * as Schema from "effect/Schema";
import type * as Scope from "effect/Scope";
import * as Stream from "effect/Stream";
import * as ChildProcessSpawner from "effect/process/ChildProcessSpawner";
import * as EffectAcpErrors from "effect-acp/errors";

import * as AcpSessionRuntime from "@t3tools/provider-acp/server/AcpSessionRuntime";
import type { ProviderReplayGate } from "@t3tools/provider-testing/replayGate";
import { ACP_PROTOCOL, type AcpAdapterV2RuntimeInput } from "@t3tools/provider-acp/server/adapter";

export const AcpReplayTranscript = Schema.Struct({
  provider: ProviderDriverKind,
  protocol: Schema.Literal(ACP_PROTOCOL),
  version: Schema.String,
  scenario: Schema.String,
  metadata: Schema.optional(Schema.Record(Schema.String, Schema.Unknown)),
  entries: Schema.Array(ProviderReplayEntry),
});
export type AcpReplayTranscript = typeof AcpReplayTranscript.Type;

const decodeAcpReplayTranscriptSchema = Schema.decodeUnknownEffect(AcpReplayTranscript);
const encodeReplayTranscriptJson = Schema.encodeSync(Schema.fromJsonString(Schema.Unknown));

export class AcpReplayTranscriptDecodeError extends Schema.TaggedError<AcpReplayTranscriptDecodeError>()(
  "AcpReplayTranscriptDecodeError",
  {
    expectedProvider: ProviderDriverKind,
    driver: Schema.optional(Schema.String),
    protocol: Schema.optional(Schema.String),
    scenario: Schema.optional(Schema.String),
    cause: Schema.Defect(),
  },
) {
  override get message(): string {
    return `Failed to decode ${this.expectedProvider} ACP replay transcript for scenario ${this.scenario ?? "<unknown>"}.`;
  }
}

const isAcpReplayTranscriptDecodeError = Schema.is(AcpReplayTranscriptDecodeError);

function metadataFromTranscript(transcript: ProviderReplayTranscript): {
  readonly provider?: string;
  readonly protocol?: string;
  readonly scenario?: string;
} {
  return {
    provider: transcript.provider,
    protocol: transcript.protocol,
    scenario: transcript.scenario,
  };
}

export function decodeAcpReplayTranscript(
  transcript: ProviderReplayTranscript,
  expectedProvider: ProviderDriverKind,
  options: { readonly retargetProvider?: boolean } = {},
): Effect.Effect<AcpReplayTranscript, AcpReplayTranscriptDecodeError> {
  const candidate = options.retargetProvider
    ? { ...transcript, provider: expectedProvider }
    : transcript;
  return decodeAcpReplayTranscriptSchema(candidate).pipe(
    Effect.filterOrFail(
      (decoded) => decoded.provider === expectedProvider,
      () =>
        new AcpReplayTranscriptDecodeError({
          expectedProvider,
          ...metadataFromTranscript(transcript),
          cause: `Expected provider ${expectedProvider}, received ${transcript.provider}`,
        }),
    ),
    Effect.mapError((cause) =>
      isAcpReplayTranscriptDecodeError(cause)
        ? cause
        : new AcpReplayTranscriptDecodeError({
            expectedProvider,
            ...metadataFromTranscript(transcript),
            cause,
          }),
    ),
  );
}

interface ReplayStatus {
  readonly cursor: number;
  readonly total: number;
  readonly failure?: unknown;
}

function decodeReplayStatus(raw: string): ReplayStatus {
  const value: unknown = JSON.parse(raw);
  if (
    typeof value !== "object" ||
    value === null ||
    typeof Reflect.get(value, "cursor") !== "number" ||
    typeof Reflect.get(value, "total") !== "number"
  ) {
    throw new Error("ACP replay status is malformed.");
  }
  return value as ReplayStatus;
}

export function makeAcpReplayCompletenessAssertion(
  fileSystem: FileSystem.FileSystem,
  statusPath: string,
  transcript: AcpReplayTranscript,
): Effect.Effect<void, EffectAcpErrors.AcpError> {
  return fileSystem.readFileString(statusPath).pipe(
    Effect.mapError(
      (cause) =>
        new EffectAcpErrors.AcpTransportError({
          detail: `Failed to read ACP replay status for ${transcript.scenario}`,
          cause,
        }),
    ),
    Effect.flatMap((raw) =>
      Effect.try({
        try: () => decodeReplayStatus(raw),
        catch: (cause) =>
          new EffectAcpErrors.AcpTransportError({
            detail: `Failed to decode ACP replay status for ${transcript.scenario}`,
            cause,
          }),
      }),
    ),
    Effect.flatMap((status) => {
      if (
        status.failure === undefined &&
        status.cursor === transcript.entries.length &&
        status.total === transcript.entries.length
      ) {
        return Effect.void;
      }
      return Effect.fail(
        new EffectAcpErrors.AcpTransportError({
          detail: `ACP replay did not consume all frames for ${transcript.scenario}`,
          cause: status,
        }),
      );
    }),
  );
}

function acpReplayAgentArgs(scriptPath: string): ReadonlyArray<string> {
  return ["--experimental-strip-types", scriptPath];
}

const isReplaySessionUpdateFrame = Schema.is(
  Schema.Struct({
    kind: Schema.Literal("notification"),
    method: Schema.Literal("session/update"),
    params: Schema.Struct({
      sessionId: Schema.String,
      _meta: Schema.Struct({ eventId: Schema.String }),
    }),
  }),
);

type SessionUpdateHandler = Parameters<
  AcpSessionRuntime.AcpSessionRuntime["Service"]["handleSessionUpdate"]
>[0];

/** A native-handler receipt fences synthetic time; the scenario's final run checks fence persistence. */
export function withReplaySessionUpdateReceipts(
  transcript: AcpReplayTranscript,
  replayGate: ProviderReplayGate,
  handler: SessionUpdateHandler,
): SessionUpdateHandler {
  const labelsBySession = new Map<string, Map<string, string>>();
  for (const label of replayGate.handledLabels) {
    const entries = transcript.entries.filter(
      (entry) => entry.type === "emit_inbound" && entry.label === label,
    );
    const entry = entries[0];
    if (
      entries.length !== 1 ||
      entry?.type !== "emit_inbound" ||
      !isReplaySessionUpdateFrame(entry.frame)
    ) {
      throw new Error(
        `Replay handled label ${label} must identify exactly one session/update frame.`,
      );
    }
    const {
      sessionId,
      _meta: { eventId },
    } = entry.frame.params;
    if (sessionId.trim().length === 0 || eventId.trim().length === 0) {
      throw new Error(`Replay handled label ${label} requires a sessionId and eventId.`);
    }
    const identityCount = transcript.entries.filter(
      (candidate) =>
        candidate.type === "emit_inbound" &&
        isReplaySessionUpdateFrame(candidate.frame) &&
        candidate.frame.params.sessionId === sessionId &&
        candidate.frame.params._meta.eventId === eventId,
    ).length;
    if (identityCount !== 1)
      throw new Error(`Duplicate replay handled event ${sessionId}:${eventId}.`);
    const labels = labelsBySession.get(sessionId) ?? new Map<string, string>();
    labels.set(eventId, label);
    labelsBySession.set(sessionId, labels);
  }
  return (notification) => {
    const eventId = notification._meta?.eventId;
    const label =
      typeof eventId === "string"
        ? labelsBySession.get(notification.sessionId)?.get(eventId)
        : undefined;
    return handler(notification).pipe(
      Effect.tap(() =>
        label === undefined ? Effect.void : Effect.sync(() => replayGate.recordHandled(label)),
      ),
    );
  };
}

function holdGatedReplayLines(
  transcript: AcpReplayTranscript,
  replayGate: ProviderReplayGate,
): NonNullable<AcpSessionRuntime.AcpSessionRuntimeOptions["transformStdout"]> {
  const inboundLabels = transcript.entries.flatMap((entry) =>
    entry.type === "emit_inbound" ? [entry.label] : [],
  );
  let inboundIndex = 0;
  const encoder = new TextEncoder();
  return (stdout) =>
    stdout.pipe(
      Stream.decodeText,
      Stream.splitLines,
      Stream.mapEffect((line) => {
        const label = line.trim().length === 0 ? undefined : inboundLabels[inboundIndex++];
        return Effect.promise((signal) => replayGate.beforeEmit(label, signal)).pipe(
          Effect.as(encoder.encode(`${line}\n`)),
        );
      }),
    );
}

export function makeAcpReplayRuntime(input: {
  readonly transcript: AcpReplayTranscript;
  readonly statusPath: string;
  readonly scriptPath: string;
  readonly childProcessSpawner: ChildProcessSpawner.ChildProcessSpawner["Service"];
  readonly fileSystem: FileSystem.FileSystem;
  readonly customModels?: AcpSessionRuntime.AcpSessionRuntimeOptions["customModels"];
  readonly cancelMeta?: AcpSessionRuntime.AcpSessionRuntimeOptions["cancelMeta"];
  readonly initializeMeta?: AcpSessionRuntime.AcpSessionRuntimeOptions["initializeMeta"];
  readonly replayGate?: ProviderReplayGate;
}): (
  runtimeInput: AcpAdapterV2RuntimeInput,
) => Effect.Effect<
  AcpSessionRuntime.AcpSessionRuntime["Service"],
  EffectAcpErrors.AcpError,
  Crypto.Crypto | Scope.Scope
> {
  // Recorded transcripts outgrow the kernel's 128 KiB limit on one environment
  // variable, so the agent reads the transcript from a file beside its status.
  const transcriptPath = `${input.statusPath}.transcript.json`;
  return (runtimeInput) =>
    Effect.gen(function* () {
      yield* input.fileSystem
        .writeFileString(transcriptPath, encodeReplayTranscriptJson(input.transcript))
        .pipe(
          Effect.mapError(
            (cause) =>
              new EffectAcpErrors.AcpTransportError({
                detail: `Failed to write ACP replay transcript for ${input.transcript.scenario}`,
                cause,
              }),
          ),
        );
      const context = yield* Layer.build(
        AcpSessionRuntime.layer({
          ...runtimeInput,
          spawn: {
            command: process.execPath,
            args: acpReplayAgentArgs(input.scriptPath),
            cwd: runtimeInput.cwd,
            env: {
              ...process.env,
              T3_ACP_REPLAY_TRANSCRIPT_PATH: transcriptPath,
              T3_ACP_REPLAY_STATUS_PATH: input.statusPath,
              T3_ACP_REPLAY_WORKSPACE: runtimeInput.cwd,
            },
          },
          authMethodId: "replay",
          ...(input.customModels === undefined ? {} : { customModels: input.customModels }),
          ...(input.cancelMeta === undefined ? {} : { cancelMeta: input.cancelMeta }),
          ...(input.initializeMeta === undefined ? {} : { initializeMeta: input.initializeMeta }),
          ...(input.replayGate === undefined
            ? {}
            : { transformStdout: holdGatedReplayLines(input.transcript, input.replayGate) }),
        }).pipe(
          Layer.provide(
            Layer.succeed(ChildProcessSpawner.ChildProcessSpawner, input.childProcessSpawner),
          ),
        ),
      );
      const runtime = yield* Effect.service(AcpSessionRuntime.AcpSessionRuntime).pipe(
        Effect.provide(context),
      );
      const replayGate = input.replayGate;
      if (replayGate === undefined || replayGate.handledLabels.length === 0) return runtime;
      return {
        ...runtime,
        handleSessionUpdate: (handler) =>
          runtime.handleSessionUpdate(
            withReplaySessionUpdateReceipts(input.transcript, replayGate, handler),
          ),
      };
    });
}

export function makeTurnInput(input: {
  readonly threadId: ThreadId;
  readonly providerThread: OrchestrationV2ProviderThread;
  readonly instanceId: ProviderInstanceId;
  readonly runtimePolicy: ProviderAdapter.ProviderAdapterV2RuntimePolicy;
  readonly now: DateTime.Utc;
  readonly ordinal?: number;
  readonly modelSelection?: ModelSelection;
  readonly messageCreatedBy?: "user" | "agent";
  readonly messageCreationSource?: "web" | "mobile" | "mcp" | "provider" | "server";
  readonly messageText?: string;
}): ProviderAdapter.ProviderAdapterV2TurnInput {
  const ordinal = input.ordinal ?? 1;
  const suffix = `${input.threadId}:${ordinal}`;
  const modelSelection =
    input.modelSelection ?? ({ instanceId: input.instanceId, model: "default" } as const);
  return {
    appThread: {
      createdBy: "user",
      creationSource: "web",
      id: input.threadId,
      projectId: ProjectId.make(`project:${input.threadId}`),
      title: "ACP adapter test",
      providerInstanceId: input.instanceId,
      modelSelection,
      runtimeMode: "approval-required",
      interactionMode: "default",
      branch: null,
      worktreePath: null,
      activeProviderThreadId: input.providerThread.id,
      lineage: {
        parentThreadId: null,
        relationshipToParent: null,
        rootThreadId: input.threadId,
      },
      forkedFrom: null,
      createdAt: input.now,
      updatedAt: input.now,
      archivedAt: null,
      settledOverride: null,
      settledAt: null,
      lastVisitedAt: null,
      deletedAt: null,
    },
    threadId: input.threadId,
    runId: RunId.make(`run:${suffix}`),
    runOrdinal: ordinal,
    providerTurnOrdinal: ordinal,
    attemptId: RunAttemptId.make(`attempt:${suffix}`),
    rootNodeId: NodeId.make(`node:${suffix}`),
    providerThread: input.providerThread,
    message: {
      createdBy: input.messageCreatedBy ?? "user",
      creationSource: input.messageCreationSource ?? "web",
      messageId: MessageId.make(`message:${suffix}`),
      text: input.messageText ?? "test prompt",
      attachments: [],
    },
    modelSelection,
    runtimePolicy: input.runtimePolicy,
  };
}
