import { msg } from "gt-react";
import { m } from "~/i18n/messages";
import type {
  OfficeDocument,
  OfficeDocumentFormat,
  OfficeWorkerResponse,
} from "./officeDocumentParser";

export const OFFICE_PREVIEW_MAX_BYTES = 5 * 1024 * 1024;
export const OFFICE_PREVIEW_TRANSFER_TIMEOUT_MS = 120_000;
export const OFFICE_PREVIEW_STARTUP_TIMEOUT_MS = 120_000;
export const OFFICE_PREVIEW_TIMEOUT_MS = 15_000;

export interface OfficeDocumentSource {
  readonly name: string;
  readonly format: OfficeDocumentFormat;
  readonly src?: string | null | undefined;
  readonly file?: Blob | null | undefined;
}

export type OfficeDocumentState =
  | { readonly status: "loading" }
  | { readonly status: "success"; readonly document: OfficeDocument }
  | { readonly status: "error"; readonly message: string };

function checkSize(size: number) {
  if (size > OFFICE_PREVIEW_MAX_BYTES) {
    throw new Error(
      "Office previews are limited to 5 MiB. Save this file to open it in another app.",
    );
  }
}

/** Bound the actual bytes received, including responses without a trustworthy Content-Length. */
export async function readOfficeDocumentBytes(source: OfficeDocumentSource, signal: AbortSignal) {
  signal.throwIfAborted();
  let body: ReadableStream<Uint8Array> | null;
  if (source.file) {
    checkSize(source.file.size);
    body = source.file.stream();
  } else if (source.src) {
    const response = await fetch(source.src, { signal });
    if (!response.ok) {
      await response.body?.cancel();
      throw new Error(`Unable to load Office document (HTTP ${response.status}).`);
    }
    const declaredSize = Number(response.headers.get("Content-Length"));
    try {
      checkSize(declaredSize);
    } catch (cause) {
      await response.body?.cancel();
      throw cause;
    }
    body = response.body;
  } else {
    throw new Error("The Office document is unavailable.");
  }
  if (!body) throw new Error("The Office document is empty.");
  const reader = body.getReader();
  const abort = () => {
    void reader.cancel().catch(() => undefined);
  };
  signal.addEventListener("abort", abort, { once: true });
  // A fixed buffer also bounds bookkeeping for responses split into many tiny chunks.
  const bytes = new Uint8Array(OFFICE_PREVIEW_MAX_BYTES);
  let length = 0;
  try {
    signal.throwIfAborted();
    while (true) {
      const result = await reader.read();
      signal.throwIfAborted();
      if (result.done) break;
      const nextLength = length + result.value.byteLength;
      checkSize(nextLength);
      bytes.set(result.value, length);
      length = nextLength;
    }
  } finally {
    signal.removeEventListener("abort", abort);
    await reader.cancel().catch(() => undefined);
    reader.releaseLock();
  }
  return bytes.buffer.slice(0, length);
}

/** A subscription owns a single load and worker; releasing its last subscriber stops both. */
export function createOfficeDocumentClient(source: OfficeDocumentSource, initialRevision = 0) {
  let state: OfficeDocumentState = { status: "loading" };
  const listeners = new Set<() => void>();
  let stop: (() => void) | null = null;
  let generation = initialRevision;
  const publish = (next: OfficeDocumentState) => {
    state = next;
    listeners.forEach((listener) => listener());
  };
  const start = () => {
    const current = ++generation;
    const controller = new AbortController();
    let worker: Worker | null = null;
    let finished = false;
    const cleanup = () => {
      controller.abort();
      worker?.terminate();
      worker = null;
      clearTimeout(timeout);
    };
    const finish = (next: OfficeDocumentState) => {
      if (finished || current !== generation) return;
      finished = true;
      cleanup();
      publish(next);
    };
    let timeout = setTimeout(() => {
      finish({
        status: "error",
        message: m(
          msg("Office document download timed out. Save this file to open it in another app."),
        ),
      });
    }, OFFICE_PREVIEW_TRANSFER_TIMEOUT_MS);
    stop = () => {
      finished = true;
      generation++;
      cleanup();
      stop = null;
    };
    publish({ status: "loading" });
    void readOfficeDocumentBytes(source, controller.signal)
      .then((buffer) => {
        if (finished || current !== generation) return;
        clearTimeout(timeout);
        timeout = setTimeout(() => {
          finish({
            status: "error",
            message: m(
              msg("Office preview startup timed out. Save this file to open it in another app."),
            ),
          });
        }, OFFICE_PREVIEW_STARTUP_TIMEOUT_MS);
        worker = new Worker(new URL("./officeDocument.worker.ts", import.meta.url), {
          type: "module",
        });
        const readyWorker = worker;
        let documentPosted = false;
        worker.addEventListener("message", (event: MessageEvent<OfficeWorkerResponse>) => {
          if (finished || current !== generation) return;
          if ("type" in event.data) {
            if (event.data.type !== "ready" || documentPosted) return;
            documentPosted = true;
            clearTimeout(timeout);
            timeout = setTimeout(() => {
              finish({
                status: "error",
                message: m(
                  msg("Office preview timed out. Save this file to open it in another app."),
                ),
              });
            }, OFFICE_PREVIEW_TIMEOUT_MS);
            try {
              readyWorker.postMessage(
                { id: current, name: source.name, format: source.format, buffer },
                [buffer],
              );
            } catch (cause) {
              finish({
                status: "error",
                message:
                  cause instanceof Error
                    ? cause.message
                    : m(msg("Unable to load this Office document.")),
              });
            }
            return;
          }
          if (!documentPosted || event.data.id !== current) return;
          finish(
            event.data.ok
              ? { status: "success", document: event.data.document }
              : { status: "error", message: event.data.error },
          );
        });
        worker.addEventListener("error", () =>
          finish({
            status: "error",
            message: m(msg("Unable to read this Office document. Save it to open in another app.")),
          }),
        );
        worker.addEventListener("messageerror", () =>
          finish({ status: "error", message: m(msg("Unable to read the Office preview result.")) }),
        );
      })
      .catch((cause: unknown) => {
        if (!finished && current === generation) {
          finish({
            status: "error",
            message:
              cause instanceof Error
                ? cause.message
                : m(msg("Unable to load this Office document.")),
          });
        }
      });
  };
  return {
    getSnapshot: () => state,
    subscribe: (listener: () => void) => {
      listeners.add(listener);
      if (stop === null) start();
      return () => {
        listeners.delete(listener);
        if (listeners.size === 0) stop?.();
      };
    },
  };
}
