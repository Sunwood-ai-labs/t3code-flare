import { act, render } from "../../test/render";
import { expect, it, vi } from "vite-plus/test";

vi.mock("~/localApi", () => ({
  ensureLocalApi: () => ({
    persistence: { getClientSettings: async () => null, setClientSettings: async () => {} },
  }),
}));

import {
  ensureClientSettingsHydrated,
  persistClientSettingsPatch,
  useEnvironmentArtworkImageUrl,
} from "./useSettings";

function ArtworkReader() {
  const url = useEnvironmentArtworkImageUrl();
  return <output>{url}</output>;
}

it("updates the rendered artwork URL when its setting changes and clears it on reset", async () => {
  await ensureClientSettingsHydrated();
  await persistClientSettingsPatch({ environmentArtworkImageUrl: "" });
  await using screen = await render(<ArtworkReader />);
  const { container } = screen;
  expect(container.querySelector("output")?.textContent).toBe("");
  await act(() =>
    persistClientSettingsPatch({ environmentArtworkImageUrl: "https://example.test/first.png" }),
  );
  expect(container.querySelector("output")?.textContent).toBe("https://example.test/first.png");
  await act(() =>
    persistClientSettingsPatch({ environmentArtworkImageUrl: "https://example.test/second.png" }),
  );
  expect(container.querySelector("output")?.textContent).toBe("https://example.test/second.png");
  await act(() => persistClientSettingsPatch({ environmentArtworkImageUrl: "" }));
  expect(container.querySelector("output")?.textContent).toBe("");
});
