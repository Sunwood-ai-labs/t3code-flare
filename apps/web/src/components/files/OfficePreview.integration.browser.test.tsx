import { EnvironmentId, ThreadId } from "@t3tools/contracts";
import { act, render, stubGlobal } from "../../../test/render";
import * as XLSX from "xlsx";
import { describe, expect, it, vi } from "vite-plus/test";

import { AttachmentFilePreview } from "./AttachmentFilePreview";
import { WorkspaceOfficePreview } from "./WorkspaceOfficePreview";
import {
  parseOfficeDocument,
  type OfficeWorkerRequest,
  type OfficeWorkerResponse,
} from "./officeDocumentParser";

const mocks = vi.hoisted(() => ({
  url: "https://environment.test/api/assets/signed/book.xlsx",
  refresh: vi.fn<() => Promise<string | null>>(),
}));
vi.mock(import("~/assets/assetUrls"), async (original) => ({
  ...(await original()),
  useAssetUrlRefresh: () => mocks.refresh,
  useAssetUrlState: () => ({ _tag: "Success", url: mocks.url, expiresAt: Number.MAX_SAFE_INTEGER }),
}));
function workbookBytes(bookType: "xlsx" | "biff8" = "xlsx", value = "ねこ / real workbook") {
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(
    workbook,
    XLSX.utils.aoa_to_sheet([
      ["Product", "Total"],
      [value, 1800],
    ]),
    "売上",
  );
  return new Uint8Array(XLSX.write(workbook, { type: "array", bookType, compression: true }));
}

const environmentId = EnvironmentId.make("test-environment");
const threadRef = { environmentId, threadId: ThreadId.make("test-thread") };

function setup() {
  mocks.url = "https://environment.test/api/assets/signed/book.xlsx";
  mocks.refresh.mockReset().mockImplementation(async () => mocks.url);
  const instances: ParserWorker[] = [];
  let posted: ((worker: ParserWorker) => void) | undefined;
  /** Only the browser worker transport is replaced; its request uses the real decoder. */
  class ParserWorker extends EventTarget {
    request: OfficeWorkerRequest | undefined;
    terminate = vi.fn();
    constructor() {
      super();
      instances.push(this);
      queueMicrotask(() =>
        this.dispatchEvent(
          new MessageEvent<OfficeWorkerResponse>("message", { data: { type: "ready" } }),
        ),
      );
    }
    postMessage(request: OfficeWorkerRequest) {
      this.request = request;
      posted?.(this);
    }
    finish() {
      const request = this.request;
      if (!request) throw new Error("Worker has no document request");
      const document = parseOfficeDocument(new Uint8Array(request.buffer), request.format);
      this.dispatchEvent(
        new MessageEvent<OfficeWorkerResponse>("message", {
          data: { id: request.id, ok: true, document },
        }),
      );
    }
    fail() {
      this.dispatchEvent(new Event("error"));
    }
  }
  const worker = stubGlobal("Worker", ParserWorker);
  const createObjectUrl = vi.spyOn(URL, "createObjectURL").mockReturnValue("blob:office-local");
  const revokeObjectUrl = vi.spyOn(URL, "revokeObjectURL").mockImplementation(() => undefined);
  const fetch = stubGlobal(
    "fetch",
    vi.fn(async () => new Response(workbookBytes())),
  );
  return {
    instances,
    nextWorker: () =>
      new Promise<ParserWorker>((resolve) => {
        posted = resolve;
      }),
    [Symbol.dispose]() {
      worker[Symbol.dispose]();
      fetch[Symbol.dispose]();
      createObjectUrl.mockRestore();
      revokeObjectUrl.mockRestore();
    },
  };
}

const workspace = (
  name = "/workspace/book.xlsx",
  reloadRevision = 0,
  workspaceMutationId: string | null = null,
) => (
  <WorkspaceOfficePreview
    environmentId={environmentId}
    threadRef={threadRef}
    draft={false}
    absolutePath={name}
    workspaceRoot="/workspace"
    format="xlsx"
    workspaceMutationId={workspaceMutationId}
    reloadRevision={reloadRevision}
  />
);

function pendingAuthorization() {
  let complete: ((url: string | null) => void) | undefined;
  const promise = new Promise<string | null>((resolve) => {
    complete = resolve;
  });
  mocks.refresh.mockReturnValueOnce(promise);
  return { promise, complete: (url: string | null) => complete?.(url) };
}

const remoteAttachment = (attachmentId = "captured-xlsx") => (
  <AttachmentFilePreview
    name={`${attachmentId}.xlsx`}
    mimeType="text/plain"
    sizeBytes={100}
    asset={{ environmentId, attachmentId }}
  />
);

describe("Office entry points with the real document client and parser", () => {
  it("waits for reauthorization before starting one download per open for a cached completed mutation", async () => {
    using _setup = setup();
    await using screen = await render(null);
    const { container } = screen;
    const firstAuthorization = pendingAuthorization();
    const firstPosted = _setup.nextWorker();
    await act(() => screen.rerender(workspace("/workspace/book.xlsx", 0, "mutation-a")));
    expect(mocks.refresh).toHaveBeenCalledOnce();
    expect(fetch).not.toHaveBeenCalled();
    expect(_setup.instances).toHaveLength(0);
    await act(async () => {
      firstAuthorization.complete("https://environment.test/authorized-first.xlsx");
      await firstAuthorization.promise;
    });
    const first = await firstPosted;
    await act(() => first.finish());
    await act(() => screen.rerender(null));
    const secondAuthorization = pendingAuthorization();
    const secondPosted = _setup.nextWorker();
    await act(() => screen.rerender(workspace("/workspace/book.xlsx", 0, "mutation-a")));
    expect(mocks.refresh).toHaveBeenCalledTimes(2);
    expect(fetch).toHaveBeenCalledOnce();
    expect(_setup.instances).toHaveLength(1);
    await act(async () => {
      secondAuthorization.complete("https://environment.test/authorized-second.xlsx");
      await secondAuthorization.promise;
    });
    const second = await secondPosted;
    await act(() => second.finish());
    expect(fetch).toHaveBeenCalledTimes(2);
    expect(vi.mocked(fetch).mock.calls.map(([url]) => String(url))).toEqual([
      "https://environment.test/authorized-first.xlsx?workspace-revision=mutation-a%3A0",
      "https://environment.test/authorized-second.xlsx?workspace-revision=mutation-a%3A0",
    ]);
    expect(first.terminate).toHaveBeenCalledOnce();
    expect(second.terminate).toHaveBeenCalledOnce();
    expect(container.querySelector("tbody")?.textContent).toContain("ねこ / real workbook");
  });

  it("gates new workspace mutations and manual reloads until each authorization completes", async () => {
    using _setup = setup();
    await using screen = await render(null);
    const { container } = screen;
    const initialPosted = _setup.nextWorker();
    await act(() => screen.rerender(workspace()));
    await act(async () => {
      (await initialPosted).finish();
    });
    const mutationAuthorization = pendingAuthorization();
    const changedPosted = _setup.nextWorker();
    await act(() => screen.rerender(workspace("/workspace/book.xlsx", 0, "mutation-new")));
    expect(fetch).toHaveBeenCalledOnce();
    expect(_setup.instances).toHaveLength(1);
    expect(container.querySelector("table")).toBeNull();
    vi.mocked(fetch).mockImplementation(
      async () => new Response(workbookBytes("xlsx", "mutation value")),
    );
    await act(async () => {
      mutationAuthorization.complete(mocks.url);
      await mutationAuthorization.promise;
    });
    await act(async () => {
      (await changedPosted).finish();
    });
    expect(container.querySelector("tbody")?.textContent).toContain("mutation value");
    const manualAuthorization = pendingAuthorization();
    const manualPosted = _setup.nextWorker();
    await act(() => screen.rerender(workspace("/workspace/book.xlsx", 1, "mutation-new")));
    expect(fetch).toHaveBeenCalledTimes(2);
    expect(_setup.instances).toHaveLength(2);
    vi.mocked(fetch).mockImplementation(
      async () => new Response(workbookBytes("xlsx", "manual value")),
    );
    await act(async () => {
      manualAuthorization.complete(mocks.url);
      await manualAuthorization.promise;
    });
    await act(async () => {
      (await manualPosted).finish();
    });
    expect(container.querySelector("tbody")?.textContent).toContain("manual value");
    expect(mocks.refresh).toHaveBeenCalledTimes(2);
    expect(fetch).toHaveBeenCalledTimes(3);
    expect(_setup.instances).toHaveLength(3);
  });

  it("reauthorizes a failed preview before retrying its reader", async () => {
    using _setup = setup();
    await using screen = await render(null);
    const { container } = screen;
    const initialPosted = _setup.nextWorker();
    await act(() => screen.rerender(workspace()));
    const first = await initialPosted;
    await act(() => first.fail());
    expect(container.querySelector('[role="alert"]')).not.toBeNull();
    const retryAuthorization = pendingAuthorization();
    const retryPosted = _setup.nextWorker();
    await act(() => container.querySelector<HTMLButtonElement>('[role="alert"] button')!.click());
    expect(mocks.refresh).toHaveBeenCalledOnce();
    expect(fetch).toHaveBeenCalledOnce();
    expect(_setup.instances).toHaveLength(1);
    await act(async () => {
      retryAuthorization.complete("https://environment.test/renewed.xlsx");
      await retryAuthorization.promise;
    });
    await act(async () => {
      (await retryPosted).finish();
    });
    expect(vi.mocked(fetch).mock.calls.map(([url]) => String(url))).toEqual([
      mocks.url,
      "https://environment.test/renewed.xlsx",
    ]);
    expect(_setup.instances).toHaveLength(2);
    expect(container.querySelector("tbody")?.textContent).toContain("ねこ / real workbook");
  });

  it("ignores an older resource authorization that resolves after the current preview", async () => {
    using _setup = setup();
    await using screen = await render(null);
    const { container } = screen;
    const oldAuthorization = pendingAuthorization();
    await act(() => screen.rerender(workspace("/workspace/old.xlsx", 0, "mutation-a")));
    const currentAuthorization = pendingAuthorization();
    const posted = _setup.nextWorker();
    await act(() => screen.rerender(workspace("/workspace/current.xlsx", 0, "mutation-a")));
    await act(async () => {
      currentAuthorization.complete("https://environment.test/current.xlsx");
      await currentAuthorization.promise;
    });
    await act(async () => {
      (await posted).finish();
    });
    await act(async () => {
      oldAuthorization.complete("https://environment.test/old.xlsx");
      await oldAuthorization.promise;
    });
    expect(fetch).toHaveBeenCalledExactlyOnceWith(
      "https://environment.test/current.xlsx?workspace-revision=mutation-a%3A0",
      { signal: expect.any(AbortSignal) },
    );
    expect(_setup.instances).toHaveLength(1);
    expect(container.querySelector("tbody")?.textContent).toContain("ねこ / real workbook");
  });

  it("waits for a new authorization when returning to a previously authorized resource", async () => {
    using _setup = setup();
    await using screen = await render(null);
    const { container } = screen;
    const firstPosted = _setup.nextWorker();
    await act(() => screen.rerender(workspace("/workspace/a.xlsx", 0, "mutation-a")));
    const first = await firstPosted;
    await act(() => first.finish());
    const otherAuthorization = pendingAuthorization();
    await act(() => screen.rerender(workspace("/workspace/b.xlsx", 0, "mutation-a")));
    const returnAuthorization = pendingAuthorization();
    const returnedPosted = _setup.nextWorker();
    await act(() => screen.rerender(workspace("/workspace/a.xlsx", 0, "mutation-a")));
    expect(fetch).toHaveBeenCalledOnce();
    expect(_setup.instances).toHaveLength(1);
    expect(container.querySelector("table")).toBeNull();
    await act(async () => {
      otherAuthorization.complete("https://environment.test/stale-b.xlsx");
      await otherAuthorization.promise;
    });
    expect(fetch).toHaveBeenCalledOnce();
    await act(async () => {
      returnAuthorization.complete("https://environment.test/returned-a.xlsx");
      await returnAuthorization.promise;
    });
    const returned = await returnedPosted;
    await act(() => returned.finish());
    expect(fetch).toHaveBeenCalledTimes(2);
    expect(_setup.instances).toHaveLength(2);
    expect(vi.mocked(fetch).mock.calls[1]?.[0]).toBe(
      "https://environment.test/returned-a.xlsx?workspace-revision=mutation-a%3A0",
    );
    expect(container.querySelector("tbody")?.textContent).toContain("ねこ / real workbook");
  });

  it("recovers from authorization failure and ignores an authorization completed after closing", async () => {
    using _setup = setup();
    await using screen = await render(null);
    const { container } = screen;
    const failedAuthorization = pendingAuthorization();
    await act(() => screen.rerender(workspace("/workspace/book.xlsx", 0, "mutation-a")));
    await act(async () => {
      failedAuthorization.complete(null);
      await failedAuthorization.promise;
    });
    expect(container.querySelector('[role="alert"]')).not.toBeNull();
    expect(fetch).not.toHaveBeenCalled();
    const retryAuthorization = pendingAuthorization();
    const retryPosted = _setup.nextWorker();
    await act(() => container.querySelector<HTMLButtonElement>('[role="alert"] button')!.click());
    expect(fetch).not.toHaveBeenCalled();
    await act(async () => {
      retryAuthorization.complete("https://environment.test/recovered.xlsx");
      await retryAuthorization.promise;
    });
    const retry = await retryPosted;
    await act(() => retry.finish());
    expect(container.querySelector("tbody")?.textContent).toContain("ねこ / real workbook");
    const closingAuthorization = pendingAuthorization();
    await act(() => screen.rerender(workspace("/workspace/book.xlsx", 1, "mutation-a")));
    await act(() => screen.rerender(null));
    await act(async () => {
      closingAuthorization.complete("https://environment.test/closed.xlsx");
      await closingAuthorization.promise;
    });
    expect(fetch).toHaveBeenCalledOnce();
    expect(_setup.instances).toHaveLength(1);
    expect(container.textContent).toBe("");
  });

  it("reloads changed bytes on a manual revision even when the signed URL is unchanged", async () => {
    using _setup = setup();
    await using screen = await render(null);
    const { container } = screen;
    const firstPosted = _setup.nextWorker();
    await act(() => screen.rerender(workspace()));
    const first = await firstPosted;
    await act(() => first.finish());
    expect(container.querySelector("tbody")?.textContent).toContain("ねこ / real workbook");
    vi.mocked(fetch).mockImplementation(
      async () => new Response(workbookBytes("xlsx", "updated value")),
    );
    const secondPosted = _setup.nextWorker();
    await act(() => screen.rerender(workspace("/workspace/book.xlsx", 1)));
    const second = await secondPosted;
    await act(() => second.finish());
    expect(container.querySelector("tbody")?.textContent).toContain("updated value");
    expect(container.querySelector("tbody")?.textContent).not.toContain("ねこ / real workbook");
    expect(vi.mocked(fetch).mock.calls.map(([url]) => String(url))).toEqual([
      mocks.url,
      `${mocks.url}?workspace-revision=%3A1`,
    ]);
  });

  it("loads a signed workspace asset and displays decoded cells", async () => {
    using _setup = setup();
    await using screen = await render(null);
    const { container } = screen;
    const posted = _setup.nextWorker();
    await act(() => screen.rerender(workspace()));
    const worker = await posted;
    await act(() => worker.finish());
    expect(fetch).toHaveBeenCalledExactlyOnceWith(mocks.url, { signal: expect.any(AbortSignal) });
    expect(container.querySelector("table")?.getAttribute("aria-label")).toBe("売上 cell values");
    expect(container.querySelector("tbody")?.textContent).toContain("ねこ / real workbook1800");
    expect(container.textContent).toContain(
      "Values only; styles, charts, and formula recalculation",
    );
    expect(worker.terminate).toHaveBeenCalledOnce();
  });

  it("routes a text/plain xlsx attachment through signed binary loading rather than UTF-8 source", async () => {
    using _setup = setup();
    await using screen = await render(null);
    const { container } = screen;
    const posted = _setup.nextWorker();
    await act(() =>
      screen.rerender(
        <AttachmentFilePreview
          name="captured.xlsx"
          mimeType="text/plain"
          sizeBytes={100}
          asset={{ environmentId, attachmentId: "captured-xlsx" }}
        />,
      ),
    );
    const worker = await posted;
    await act(() => worker.finish());
    expect(mocks.refresh).toHaveBeenCalledOnce();
    expect(fetch).toHaveBeenCalledExactlyOnceWith(mocks.url, { signal: expect.any(AbortSignal) });
    expect(container.querySelector("td")?.textContent).toBe("Product");
    expect(container.querySelector("tbody")?.textContent).toContain("ねこ / real workbook1800");
    expect(container.querySelector("pre")).toBeNull();
    expect(container.querySelector('[role="alert"]')).toBeNull();
  });

  it("starts one attachment reader after retry authorization even when the signed URL is unchanged", async () => {
    using _setup = setup();
    await using screen = await render(null);
    const { container } = screen;
    const initialAuthorization = pendingAuthorization();
    const initialPosted = _setup.nextWorker();
    await act(() => screen.rerender(remoteAttachment()));
    expect(fetch).not.toHaveBeenCalled();
    expect(_setup.instances).toHaveLength(0);
    await act(async () => {
      initialAuthorization.complete(mocks.url);
      await initialAuthorization.promise;
    });
    const first = await initialPosted;
    await act(() => first.fail());
    const retryAuthorization = pendingAuthorization();
    const retryPosted = _setup.nextWorker();
    vi.mocked(fetch).mockImplementation(
      async () => new Response(workbookBytes("xlsx", "renewed attachment")),
    );
    await act(() => container.querySelector<HTMLButtonElement>('[role="alert"] button')!.click());
    expect(mocks.refresh).toHaveBeenCalledTimes(2);
    expect(fetch).toHaveBeenCalledOnce();
    expect(_setup.instances).toHaveLength(1);
    expect(container.querySelector("table")).toBeNull();
    await act(async () => {
      retryAuthorization.complete(mocks.url);
      await retryAuthorization.promise;
    });
    const retry = await retryPosted;
    await act(() => retry.finish());
    expect(fetch).toHaveBeenCalledTimes(2);
    expect(vi.mocked(fetch).mock.calls.map(([url]) => String(url))).toEqual([mocks.url, mocks.url]);
    expect(_setup.instances).toHaveLength(2);
    expect(first.terminate).toHaveBeenCalledOnce();
    expect(retry.terminate).toHaveBeenCalledOnce();
    expect(container.querySelector("tbody")?.textContent).toContain("renewed attachment");
  });

  it("does not reuse attachment authorization across A-to-B-to-A or accept results after closing", async () => {
    using _setup = setup();
    await using screen = await render(null);
    const { container } = screen;
    const firstPosted = _setup.nextWorker();
    await act(() => screen.rerender(remoteAttachment("a")));
    const first = await firstPosted;
    await act(() => first.finish());
    const otherAuthorization = pendingAuthorization();
    await act(() => screen.rerender(remoteAttachment("b")));
    const returnedAuthorization = pendingAuthorization();
    const returnedPosted = _setup.nextWorker();
    await act(() => screen.rerender(remoteAttachment("a")));
    expect(fetch).toHaveBeenCalledOnce();
    expect(_setup.instances).toHaveLength(1);
    expect(container.querySelector("table")).toBeNull();
    await act(async () => {
      otherAuthorization.complete("https://environment.test/obsolete-b.xlsx");
      await otherAuthorization.promise;
    });
    expect(fetch).toHaveBeenCalledOnce();
    await act(async () => {
      returnedAuthorization.complete("https://environment.test/returned-a.xlsx");
      await returnedAuthorization.promise;
    });
    const returned = await returnedPosted;
    await act(() => returned.finish());
    expect(fetch).toHaveBeenCalledTimes(2);
    expect(vi.mocked(fetch).mock.calls[1]?.[0]).toBe("https://environment.test/returned-a.xlsx");
    expect(container.querySelector("tbody")?.textContent).toContain("ねこ / real workbook");
    const closingAuthorization = pendingAuthorization();
    await act(() => screen.rerender(remoteAttachment("closing")));
    await act(() => screen.rerender(null));
    await act(async () => {
      closingAuthorization.complete("https://environment.test/closed.xlsx");
      await closingAuthorization.promise;
    });
    expect(fetch).toHaveBeenCalledTimes(2);
    expect(_setup.instances).toHaveLength(2);
    expect(container.textContent).toBe("");
  });

  it("retries a local Office Blob without remote authorization or HTTP loading", async () => {
    using _setup = setup();
    await using screen = await render(null);
    const { container } = screen;
    const file = new Blob([workbookBytes()]);
    const firstPosted = _setup.nextWorker();
    await act(() =>
      screen.rerender(
        <AttachmentFilePreview
          name="local.xlsx"
          mimeType="application/octet-stream"
          sizeBytes={file.size}
          file={file}
        />,
      ),
    );
    const first = await firstPosted;
    await act(() => first.fail());
    const retryPosted = _setup.nextWorker();
    await act(() => container.querySelector<HTMLButtonElement>('[role="alert"] button')!.click());
    const retry = await retryPosted;
    await act(() => retry.finish());
    expect(mocks.refresh).not.toHaveBeenCalled();
    expect(fetch).not.toHaveBeenCalled();
    expect(_setup.instances).toHaveLength(2);
    expect(first.terminate).toHaveBeenCalledOnce();
    expect(container.querySelector("tbody")?.textContent).toContain("ねこ / real workbook");
  });

  it("reads real BIFF8 from a local Blob and releases its object URL when changing to legacy ppt fallback", async () => {
    using _setup = setup();
    await using screen = await render(null);
    const { container } = screen;
    const file = new Blob([workbookBytes("biff8")], { type: "application/vnd.ms-excel" });
    const posted = _setup.nextWorker();
    await act(() =>
      screen.rerender(
        <AttachmentFilePreview
          name="local.xls"
          mimeType="application/vnd.ms-excel"
          sizeBytes={file.size}
          file={file}
        />,
      ),
    );
    const worker = await posted;
    await act(() => worker.finish());
    expect(container.querySelector("tbody")?.textContent).toContain("ねこ / real workbook1800");
    expect(fetch).not.toHaveBeenCalled();
    expect(mocks.refresh).not.toHaveBeenCalled();
    const legacy = new Blob([new Uint8Array([0xd0, 0xcf, 0x11, 0xe0])]);
    await act(() =>
      screen.rerender(
        <AttachmentFilePreview
          name="legacy.ppt"
          mimeType="application/vnd.ms-powerpoint"
          sizeBytes={legacy.size}
          file={legacy}
        />,
      ),
    );
    expect(container.textContent).toContain("No preview for this file");
    expect(container.textContent).toContain("Save it to open in an app that supports ppt files.");
    expect(container.querySelector("table")).toBeNull();
    expect(URL.revokeObjectURL).toHaveBeenCalledWith("blob:office-local");
    expect(worker.terminate).toHaveBeenCalledOnce();
  });

  it("terminates pending work on route replacement and unmount, ignoring a stale result", async () => {
    using _setup = setup();
    await using screen = await render(null);
    const { container } = screen;
    const firstPosted = _setup.nextWorker();
    await act(() => screen.rerender(workspace("/workspace/first.xlsx")));
    const first = await firstPosted;
    mocks.url = "https://environment.test/api/assets/second/book.xlsx";
    const secondPosted = _setup.nextWorker();
    await act(() => screen.rerender(workspace("/workspace/second.xlsx")));
    const second = await secondPosted;
    expect(first.terminate).toHaveBeenCalledOnce();
    await act(() => first.finish());
    expect(container.querySelector("table")).toBeNull();
    expect(container.querySelector('[role="status"][aria-label="Loading file"]')).not.toBeNull();
    await act(() => screen.rerender(null));
    expect(second.terminate).toHaveBeenCalledOnce();
    await act(() => second.finish());
    expect(container.textContent).toBe("");
  });
});
