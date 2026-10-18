import { act } from "react";
import { createRoot, type Root } from "react-dom/client";
import { afterEach, beforeEach, describe, expect, it, vi } from "vite-plus/test";

import { resolveCustomStageArtworkUrl, SidebarStageBackdrop } from "./SidebarStageBackdrop";

let root: Root;
let container: HTMLDivElement;
const requests: HTMLImageElement[] = [];
const drawImage = vi.fn();
const getContext = vi.fn<(contextId: string) => { drawImage: typeof drawImage } | null>();
const originalGetContext = HTMLCanvasElement.prototype.getContext;

beforeEach(() => {
  vi.stubGlobal("IS_REACT_ACT_ENVIRONMENT", true);
  requests.length = 0;
  drawImage.mockReset();
  getContext.mockReset().mockReturnValue({ drawImage });
  Object.defineProperty(HTMLCanvasElement.prototype, "getContext", { value: getContext });
  vi.stubGlobal("Image", function () {
    const image = document.createElement("img");
    const getAttribute = image.getAttribute.bind(image);
    const hasAttribute = image.hasAttribute.bind(image);
    const removeAttribute = image.removeAttribute.bind(image);
    image.getAttribute = (name) => getAttribute(name === "src" ? "data-test-src" : name);
    image.hasAttribute = (name) => hasAttribute(name === "src" ? "data-test-src" : name);
    image.removeAttribute = (name) => removeAttribute(name === "src" ? "data-test-src" : name);
    Object.defineProperties(image, {
      src: { set: (value: string) => image.setAttribute("data-test-src", value) },
      naturalWidth: { value: 800, configurable: true },
      naturalHeight: { value: 240, configurable: true },
    });
    requests.push(image);
    return image;
  });
  container = document.createElement("div");
  document.body.append(container);
  root = createRoot(container);
});

afterEach(async () => {
  await act(async () => root.unmount());
  container.remove();
  Object.defineProperty(HTMLCanvasElement.prototype, "getContext", { value: originalGetContext });
  vi.restoreAllMocks();
  vi.unstubAllGlobals();
});

async function renderBackdrop(customImageUrl: string | null) {
  await act(async () => {
    root.render(<SidebarStageBackdrop customImageUrl={customImageUrl} />);
  });
}

function renderedArt() {
  return container.querySelector("[data-stage-art]")?.getAttribute("data-stage-art") ?? null;
}

function latestRequest() {
  const image = requests.at(-1);
  if (!image) throw new Error("expected an artwork image request");
  return image;
}

async function dispatchImageEvent(type: "load" | "error", image = latestRequest()) {
  await act(async () => image.dispatchEvent(new Event(type)));
}

describe("resolveCustomStageArtworkUrl", () => {
  it("accepts http(s) URLs and trims them", () => {
    expect(resolveCustomStageArtworkUrl(" https://example.com/a.png ")).toBe(
      "https://example.com/a.png",
    );
    expect(resolveCustomStageArtworkUrl("http://localhost:8080/a.png", "http:")).toBe(
      "http://localhost:8080/a.png",
    );
    expect(resolveCustomStageArtworkUrl("https://example.com/image(1).png")).toBe(
      "https://example.com/image(1).png",
    );
  });

  it.each(["https:", "t3code:", "t3code-dev:"])(
    "rejects non-loopback HTTP artwork on %s pages",
    (pageProtocol) => {
      for (const host of [
        "example.com",
        "192.168.1.10",
        "0.0.0.0",
        "localhost.example.com",
        "127.0.0.1.example.com",
        "[::2]",
      ]) {
        expect(resolveCustomStageArtworkUrl(`http://${host}/a.png`, pageProtocol)).toBeNull();
      }
      expect(resolveCustomStageArtworkUrl("https://example.com/a.png", pageProtocol)).toBe(
        "https://example.com/a.png",
      );
    },
  );

  it.each(["https:", "t3code:", "t3code-dev:"])(
    "allows trustworthy loopback HTTP artwork on %s pages",
    (pageProtocol) => {
      for (const host of [
        "localhost",
        "localhost.",
        "art.localhost",
        "art.localhost.",
        "127.0.0.1",
        "127.10.20.30",
        "127.1",
        "[::1]",
      ]) {
        const value = `http://${host}:8080/a.png`;
        expect(resolveCustomStageArtworkUrl(value, pageProtocol)).toBe(value);
      }
    },
  );

  it("allows remote HTTP artwork on local HTTP web pages", () => {
    expect(resolveCustomStageArtworkUrl("http://example.com/a.png", "http:")).toBe(
      "http://example.com/a.png",
    );
  });

  it("rejects empty values, relative paths, and other schemes", () => {
    for (const value of [
      "",
      "   ",
      "C:\\Users\\me\\a.png",
      "/home/me/a.png",
      "file:///home/me/a.png",
      "data:image/png;base64,AAAA",
      "javascript:alert(1)",
    ]) {
      expect(resolveCustomStageArtworkUrl(value)).toBeNull();
    }
  });
});

describe("SidebarStageBackdrop custom image", () => {
  it("renders the built-in artwork when no custom image is set", async () => {
    await renderBackdrop(null);
    expect(renderedArt()).toBe("flare");
    expect(requests).toHaveLength(0);
    const image = container.querySelector("img");
    if (!image) throw new Error("expected the bundled Flare header image");
    await image.decode();
    expect(image.naturalWidth).toBeGreaterThan(0);
    expect(image.naturalHeight).toBeGreaterThan(0);
  });

  it("keeps built-in artwork while loading, then draws one static frame and releases the source", async () => {
    await renderBackdrop("https://example.com/animated.webp");
    const image = latestRequest();
    expect(image.getAttribute("src")).toBe("https://example.com/animated.webp");
    expect(renderedArt()).toBe("flare");
    await dispatchImageEvent("load", image);
    expect(renderedArt()).toBe("custom");
    expect(drawImage).toHaveBeenCalledExactlyOnceWith(image, 0, 0, 800, 240);
    expect(image.hasAttribute("src")).toBe(false);
    await dispatchImageEvent("load", image);
    expect(drawImage).toHaveBeenCalledTimes(1);
  });

  it("bounds the decoded canvas dimensions", async () => {
    await renderBackdrop("https://example.com/large.png");
    const image = latestRequest();
    Object.defineProperties(image, {
      naturalWidth: { value: 8192 },
      naturalHeight: { value: 4096 },
    });
    await dispatchImageEvent("load", image);
    expect(drawImage).toHaveBeenCalledExactlyOnceWith(image, 0, 0, 2048, 1024);
  });

  it("restores the bundled artwork when a loaded custom image is cleared", async () => {
    await renderBackdrop("https://example.com/a.png");
    await dispatchImageEvent("load");
    expect(renderedArt()).toBe("custom");
    await renderBackdrop(null);
    expect(renderedArt()).toBe("flare");
    expect(container.querySelector("canvas")).toBeNull();
    expect(latestRequest().hasAttribute("src")).toBe(false);
  });

  it("falls back on image failure and retries a new URL", async () => {
    await renderBackdrop("https://example.com/missing.png");
    await dispatchImageEvent("error");
    expect(renderedArt()).toBe("flare");
    expect(latestRequest().hasAttribute("src")).toBe(false);
    await renderBackdrop("https://example.com/b.png");
    await dispatchImageEvent("load");
    expect(renderedArt()).toBe("custom");
  });

  it("keeps built-in artwork when canvas rendering is unavailable", async () => {
    drawImage.mockImplementationOnce(() => {
      throw new Error("canvas unavailable");
    });
    await renderBackdrop("https://example.com/a.png");
    await dispatchImageEvent("load");
    expect(renderedArt()).toBe("flare");
    expect(latestRequest().hasAttribute("src")).toBe(false);
  });

  it("keeps built-in artwork and releases the image when no canvas context is available", async () => {
    getContext.mockReturnValueOnce(null);
    await renderBackdrop("https://example.com/a.png");
    await dispatchImageEvent("load");
    expect(renderedArt()).toBe("flare");
    expect(drawImage).not.toHaveBeenCalled();
    expect(latestRequest().hasAttribute("src")).toBe(false);
  });

  it("ignores late load and error events from a replaced image", async () => {
    await renderBackdrop("https://example.com/a.png");
    const previous = latestRequest();
    await renderBackdrop("https://example.com/b.png");
    await dispatchImageEvent("load", previous);
    await dispatchImageEvent("error", previous);
    expect(drawImage).toHaveBeenCalledTimes(0);
    await dispatchImageEvent("load");
    expect(renderedArt()).toBe("custom");
    expect(drawImage).toHaveBeenCalledTimes(1);
  });

  it("retries a failed URL after the custom image setting is cleared", async () => {
    await renderBackdrop("https://example.com/a.png");
    const first = latestRequest();
    await dispatchImageEvent("error");
    await renderBackdrop(null);
    await renderBackdrop("https://example.com/a.png");
    expect(latestRequest()).not.toBe(first);
    await dispatchImageEvent("load");
    expect(renderedArt()).toBe("custom");
  });

  it("retries a failed URL after switching to another image and back", async () => {
    await renderBackdrop("https://example.com/a.png");
    const first = latestRequest();
    await dispatchImageEvent("error");
    await renderBackdrop("https://example.com/b.png");
    await renderBackdrop("https://example.com/a.png");
    expect(latestRequest()).not.toBe(first);
    await dispatchImageEvent("load");
    expect(renderedArt()).toBe("custom");
  });
});
