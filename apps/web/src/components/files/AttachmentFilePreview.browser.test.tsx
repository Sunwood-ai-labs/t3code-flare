import { EnvironmentId } from "@t3tools/contracts";
import { act } from "react";
import * as XLSX from "xlsx";
import { createRoot, type Root } from "react-dom/client";
import { afterEach, beforeEach, describe, expect, it, vi } from "vite-plus/test";

const { refresh } = vi.hoisted(() => ({ refresh: vi.fn<() => Promise<string | null>>() }));

vi.mock(import("~/assets/assetUrls"), async (original) => ({
  ...(await original()),
  useAssetUrlRefresh: () => refresh,
}));
// The real source view highlights through a worker pool; the text it receives is what matters here.
// Never load the original: its highlighter graph can outlast waitFor on a cold CI browser.
vi.mock(import("./ReadOnlySourcePreview"), () => ({
  default: ({ text }: { text: string }) => <pre>{text}</pre>,
}));

import { AttachmentFilePreview } from "./AttachmentFilePreview";

describe("attachment HTML preview recovery", () => {
  const originalUrl = "data:text/html,original";
  const renewedUrl = "data:text/html,renewed";
  let now = 0;
  let root: Root;
  let container: HTMLDivElement;

  beforeEach(async () => {
    await import("./ReadOnlySourcePreview");
    vi.stubGlobal("IS_REACT_ACT_ENVIRONMENT", true);
    now = 0;
    vi.spyOn(Date, "now").mockImplementation(() => now);
    refresh.mockReset().mockResolvedValueOnce(originalUrl).mockResolvedValue(renewedUrl);
    vi.stubGlobal(
      "fetch",
      vi.fn(async () => new Response("<p>Captured HTML</p>")),
    );
    container = document.createElement("div");
    document.body.append(container);
    root = createRoot(container);
  });

  afterEach(async () => {
    await act(async () => root.unmount());
    container.remove();
    vi.restoreAllMocks();
    vi.unstubAllGlobals();
  });

  const openRemote = () =>
    act(async () =>
      root.render(
        <AttachmentFilePreview
          name="document.html"
          mimeType="text/html"
          sizeBytes={100}
          asset={{ environmentId: EnvironmentId.make("test-environment"), attachmentId: "html" }}
        />,
      ),
    );

  const toggleMode = async (label: string) => {
    const button = container.querySelector<HTMLButtonElement>(`button[aria-label="${label}"]`);
    expect(button, label).not.toBeNull();
    await act(async () => button!.click());
  };
  const alert = () => container.querySelector('[role="alert"]');

  it("keeps a renewed URL for subsequent rendered and source views", async () => {
    await openRemote();
    expect(container.querySelector("iframe")?.getAttribute("src")).toBe(originalUrl);
    now = 61 * 60_000;
    await toggleMode("Show HTML source");
    await act(async () => {
      await import("./ReadOnlySourcePreview");
    });
    expect(fetch).toHaveBeenCalledExactlyOnceWith(renewedUrl, expect.any(Object));
    await vi.waitFor(() =>
      expect(container.querySelector("pre")?.textContent).toBe("<p>Captured HTML</p>"),
    );

    await toggleMode("Show rendered page");
    expect(container.querySelector("iframe")?.getAttribute("src")).toBe(renewedUrl);
    await toggleMode("Show HTML source");
    expect(refresh).toHaveBeenCalledTimes(2);
    expect(vi.mocked(fetch).mock.calls.map(([url]) => url)).toEqual([renewedUrl, renewedUrl]);
  });

  it("does not fetch or mark an expired URL fresh when reauthorization is unavailable", async () => {
    await openRemote();
    now = 61 * 60_000;
    refresh.mockResolvedValue(null);
    await toggleMode("Show HTML source");
    expect(fetch).not.toHaveBeenCalled();
    expect(alert()?.textContent).toContain("Reconnect to the environment and try again.");

    await toggleMode("Show rendered page");
    await toggleMode("Show HTML source");
    expect(refresh).toHaveBeenCalledTimes(3);
    expect(fetch).not.toHaveBeenCalled();
  });

  it("can return to rendered HTML after local source decoding fails", async () => {
    const file = new Blob([new Uint8Array([0x3c, 0x70, 0x3e, 0xe9])], { type: "text/html" });
    await act(async () =>
      root.render(
        <AttachmentFilePreview
          name="document.html"
          mimeType="text/html"
          sizeBytes={file.size}
          file={file}
        />,
      ),
    );
    await toggleMode("Show HTML source");
    await vi.waitFor(() => expect(alert()?.textContent).toContain("not UTF-8"));

    await toggleMode("Show rendered page");
    expect(alert()).toBeNull();
    expect(container.querySelector("iframe")?.title).toBe("document.html");
  });

  it("renders a real xlsx attachment with the legacy Excel MIME as Unicode cells and cached formulas", async () => {
    const workbook = XLSX.utils.book_new();
    const sheet = XLSX.utils.aoa_to_sheet([
      ["商品 🐈", "数量", "単価", "合計"],
      ["りんご café", 12, 150, 1800],
    ]);
    sheet.D2 = { t: "n", f: "B2*C2", v: 1800 };
    XLSX.utils.book_append_sheet(workbook, sheet, "売上");
    const file = new Blob(
      [XLSX.write(workbook, { type: "array", bookType: "xlsx", bookSST: true })],
      { type: "application/vnd.ms-excel" },
    );
    await act(async () =>
      root.render(
        <AttachmentFilePreview
          name="report.xlsx"
          mimeType={file.type}
          sizeBytes={file.size}
          file={file}
        />,
      ),
    );
    await vi.waitFor(
      () =>
        expect(container.querySelector("table")?.getAttribute("aria-label")).toBe(
          "売上 cell values",
        ),
      { timeout: 10_000 },
    );
    expect([...container.querySelectorAll("td")].map((cell) => cell.textContent)).toEqual([
      "商品 🐈",
      "数量",
      "単価",
      "合計",
      "りんご café",
      "12",
      "150",
      "1800",
    ]);
    expect(alert()).toBeNull();
    expect(container.querySelector("pre")).toBeNull();
    expect(fetch).not.toHaveBeenCalled();
    expect(refresh).not.toHaveBeenCalled();
  });
});
