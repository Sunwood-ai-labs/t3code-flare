import { act, render, stubGlobal } from "../../../test/render";
import type { ReactNode } from "react";
import { describe, expect, it, vi } from "vite-plus/test";
import { userEvent } from "vite-plus/test/browser";

import OfficeDocumentPreview from "./OfficeDocumentPreview";
import type { OfficeDocumentSource, OfficeDocumentState } from "./officeDocumentClient";
import type { OfficeWorkerRequest } from "./officeDocumentParser";

const { store, createClient } = vi.hoisted(() => ({
  store: {
    state: { status: "loading" } as OfficeDocumentState,
    subscribe: vi.fn(() => () => undefined),
  },
  createClient: vi.fn(),
}));
vi.mock("./officeDocumentClient", () => ({ createOfficeDocumentClient: createClient }));
vi.mock("~/components/ui/button", () => ({
  Button: (props: Record<string, unknown>) => <button {...props} />,
}));
vi.mock("./fileSurfaceChrome", () => ({
  FileSurfaceFailure: ({ message, onRetry }: { message: string; onRetry?: () => void }) => (
    <div role="alert">
      {message}
      <button onClick={onRetry}>Try again</button>
    </div>
  ),
  FileSurfaceLoading: () => <div role="status">Loading</div>,
  FileSurfaceNotice: ({ children }: { children: ReactNode }) => <div role="status">{children}</div>,
}));

function setup() {
  store.state = { status: "loading" };
  store.subscribe.mockReset().mockImplementation(() => () => undefined);
  createClient
    .mockReset()
    .mockImplementation(() => ({ subscribe: store.subscribe, getSnapshot: () => store.state }));
}

describe("Office document controls", () => {
  it("keeps all 20 slide choices focusable and displays the last selected slide", async () => {
    setup();
    await using screen = await render(null);
    const { container } = screen;
    store.state = {
      status: "success",
      document: {
        kind: "presentation",
        warnings: [],
        truncated: false,
        slides: Array.from({ length: 20 }, (_, index) => ({
          number: index + 1,
          text: [`Slide ${index + 1} body`],
          truncated: false,
          width: 9144000,
          height: 5143500,
          elements: [
            {
              kind: "text" as const,
              text: `Slide ${index + 1} body`,
              x: 100000,
              y: 100000,
              width: 8000000,
              height: 1000000,
              fontSize: 24,
            },
          ],
        })),
      },
    };
    await screen.rerender(<OfficeDocumentPreview name="20-slides.pptx" format="pptx" />);
    const buttons = [
      ...container.querySelectorAll<HTMLButtonElement>('nav[aria-label="Slides"] button'),
    ];
    expect(buttons.map((button) => button.textContent)).toEqual(
      Array.from({ length: 20 }, (_, index) => `Slide ${index + 1}`),
    );
    expect(buttons.every((button) => button.tabIndex === 0 && !button.disabled)).toBe(true);
    const last = buttons.at(-1)!;
    last.focus();
    expect(document.activeElement).toBe(last);
    await userEvent.keyboard("{Enter}");
    expect(document.activeElement).toBe(last);
    expect(last.getAttribute("aria-pressed")).toBe("true");
    expect(buttons[0]!.getAttribute("aria-pressed")).toBe("false");
    expect(container.querySelector("section")?.getAttribute("aria-label")).toBe("Slide 20 preview");
    expect(container.querySelector("section")?.textContent).toBe("Slide 20 body");
  });

  it("keeps all 20 long worksheet names accessible and lets the last focused button select its cells", async () => {
    setup();
    await using screen = await render(null);
    const { container } = screen;
    const names = Array.from({ length: 20 }, (_, index) =>
      `Worksheet_${String(index + 1).padStart(2, "0")}_`.padEnd(31, "x"),
    );
    store.state = {
      status: "success",
      document: {
        kind: "spreadsheet",
        warnings: [],
        sheets: names.map((name) => ({
          name,
          rows: [[`${name} value`]],
          totalRows: 1,
          totalColumns: 1,
          truncated: false,
        })),
      },
    };
    await screen.rerender(<OfficeDocumentPreview name="20-sheets.xlsx" format="xlsx" />);
    const buttons = [
      ...container.querySelectorAll<HTMLButtonElement>('nav[aria-label="Worksheets"] button'),
    ];
    expect(buttons.map((button) => button.textContent)).toEqual(names);
    expect(buttons.every((button) => button.tabIndex === 0 && !button.disabled)).toBe(true);
    const last = buttons.at(-1)!;
    last.focus();
    expect(document.activeElement).toBe(last);
    await userEvent.keyboard("{Enter}");
    expect(document.activeElement).toBe(last);
    expect(last.getAttribute("aria-pressed")).toBe("true");
    expect(buttons[0]!.getAttribute("aria-pressed")).toBe("false");
    expect(container.querySelector("td")?.textContent).toBe(`${names.at(-1)} value`);
    expect(container.querySelector("table")?.getAttribute("aria-label")).toBe(
      `${names.at(-1)} cell values`,
    );
  });

  it("lets a user recover text omitted from the basic slide canvas through the outline", async () => {
    setup();
    await using screen = await render(null);
    const { container } = screen;
    store.state = {
      status: "success",
      document: {
        kind: "presentation",
        warnings: ["Unsupported drawing omitted."],
        truncated: false,
        slides: [
          {
            number: 1,
            text: ["Visible title", "Text from unsupported drawing"],
            truncated: false,
            width: 9144000,
            height: 5143500,
            elements: [
              {
                kind: "text",
                text: "Visible title",
                x: 914400,
                y: 514350,
                width: 7315200,
                height: 1028700,
                fontSize: 24,
              },
            ],
          },
        ],
      },
    };
    await screen.rerender(<OfficeDocumentPreview name="deck.pptx" format="pptx" />);
    expect(container.textContent).toContain("Visible title");
    expect(container.textContent).not.toContain("Text from unsupported drawing");
    const outline = [...container.querySelectorAll("button")].find(
      (button) => button.textContent === "Text outline",
    );
    expect(outline).toBeDefined();
    await act(() => outline!.click());
    expect([...container.querySelectorAll("p")].map((paragraph) => paragraph.textContent)).toEqual([
      "Visible title",
      "Text from unsupported drawing",
    ]);
  });

  it("retries the same local Blob after a parser failure without reopening the preview", async () => {
    setup();
    let attempts = 0;
    let parsed: (() => void) | undefined;
    let posted = new Promise<void>((resolve) => {
      parsed = resolve;
    });
    const terminate = vi.fn();
    using _worker = stubGlobal(
      "Worker",
      class extends EventTarget {
        terminate = terminate;
        constructor() {
          super();
          queueMicrotask(() =>
            this.dispatchEvent(new MessageEvent("message", { data: { type: "ready" } })),
          );
        }
        postMessage(request: OfficeWorkerRequest) {
          attempts++;
          this.dispatchEvent(
            new MessageEvent("message", {
              data:
                attempts === 1
                  ? { id: request.id, ok: false, error: "Parser failed once" }
                  : {
                      id: request.id,
                      ok: true,
                      document: {
                        kind: "spreadsheet",
                        warnings: [],
                        sheets: [
                          {
                            name: "Recovered",
                            rows: [["Readable value"]],
                            totalRows: 1,
                            totalColumns: 1,
                            truncated: false,
                          },
                        ],
                      },
                    },
            }),
          );
          parsed?.();
        }
      },
    );
    await using screen = await render(null);
    const { container } = screen;
    const actual =
      await vi.importActual<typeof import("./officeDocumentClient")>("./officeDocumentClient");
    createClient.mockImplementation((source: OfficeDocumentSource, revision: number) =>
      actual.createOfficeDocumentClient(source, revision),
    );
    const file = new Blob(["test"]);
    Object.defineProperty(file, "stream", {
      value: () =>
        new ReadableStream<Uint8Array>({
          start(controller) {
            controller.enqueue(new Uint8Array([1, 2, 3]));
            controller.close();
          },
        }),
    });
    await screen.rerender(<OfficeDocumentPreview name="same.xlsx" format="xlsx" file={file} />);
    await act(async () => {
      await posted;
    });
    expect(container.querySelector('[role="alert"]')?.textContent).toContain("Parser failed once");
    posted = new Promise<void>((resolve) => {
      parsed = resolve;
    });
    await act(() => {
      container.querySelector<HTMLButtonElement>('[role="alert"] button')!.click();
    });
    await act(async () => {
      await posted;
    });
    expect(container.querySelector("td")?.textContent).toBe("Readable value");
    expect(attempts).toBe(2);
    expect(terminate).toHaveBeenCalledTimes(2);
  });
  it("changes worksheets using buttons while retaining escaped cell text", async () => {
    setup();
    await using screen = await render(null);
    const { container } = screen;
    store.state = {
      status: "success",
      document: {
        kind: "spreadsheet",
        warnings: [],
        sheets: [
          {
            name: "First",
            rows: [["<script>literal</script>"]],
            totalRows: 1,
            totalColumns: 1,
            truncated: false,
          },
          {
            name: "Second",
            rows: [["second value"]],
            totalRows: 1,
            totalColumns: 1,
            truncated: false,
          },
        ],
      },
    };
    await screen.rerender(<OfficeDocumentPreview name="book.xlsx" format="xlsx" />);
    expect(container.querySelector("td")?.textContent).toBe("<script>literal</script>");
    expect(container.querySelector("script")).toBeNull();
    await act(() => container.querySelectorAll("button")[1]!.click());
    expect(container.querySelector("td")?.textContent).toBe("second value");
    expect(
      [...container.querySelectorAll("button")].map((button) =>
        button.getAttribute("aria-pressed"),
      ),
    ).toEqual(["false", "true"]);
  });

  it("leaves typing and cursor keys alone and changes slides only through its controls", async () => {
    setup();
    await using screen = await render(null);
    const { container } = screen;
    store.state = {
      status: "success",
      document: {
        kind: "presentation",
        warnings: [],
        truncated: false,
        slides: [
          {
            number: 1,
            text: ["First slide"],
            truncated: false,
            width: 9144000,
            height: 5143500,
            elements: [],
          },
          {
            number: 2,
            text: ["Second slide"],
            truncated: false,
            width: 9144000,
            height: 5143500,
            elements: [],
          },
        ],
      },
    };
    await screen.rerender(
      <>
        <textarea aria-label="Chat input" />
        <OfficeDocumentPreview name="deck.pptx" format="pptx" />
      </>,
    );
    const input = container.querySelector("textarea")!;
    input.focus();
    for (const key of [" ", "ArrowRight", "PageDown"]) {
      input.dispatchEvent(new KeyboardEvent("keydown", { key, bubbles: true }));
    }
    expect(document.activeElement).toBe(input);
    expect(container.querySelector("section")?.getAttribute("aria-label")).toBe("Slide 1 preview");
    await act(() => container.querySelectorAll("button")[1]!.click());
    expect(container.querySelector("section")?.getAttribute("aria-label")).toBe("Slide 2 preview");
  });
});
