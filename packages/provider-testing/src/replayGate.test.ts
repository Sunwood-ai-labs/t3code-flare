import { describe, expect, it } from "vite-plus/test";

import { makeProviderReplayGate } from "./replayGate.ts";

describe("makeProviderReplayGate", () => {
  it("holds a gated emit until its label is released", async () => {
    const gate = makeProviderReplayGate(["approval"]);
    let emitted = false;
    const emit = gate.beforeEmit("approval").then(() => {
      emitted = true;
    });

    expect(await gate.waitForReached("approval")).toBe(true);
    expect(emitted).toBe(false);
    expect(gate.release("approval")).toBe(true);
    await emit;
    expect(emitted).toBe(true);
    expect(gate.release("approval")).toBe(false);
  });

  it("passes ungated labels through and stops waiting when the emit is aborted", async () => {
    const gate = makeProviderReplayGate(["approval"]);
    await gate.beforeEmit(undefined);
    await gate.beforeEmit("other");

    const controller = new AbortController();
    const emit = gate.beforeEmit("approval", controller.signal);
    controller.abort();
    await emit;
    expect(gate.hasReached("approval")).toBe(true);
  });

  it("awaits a handled receipt without turning that label into an emission gate", async () => {
    const gate = makeProviderReplayGate([], ["last-output"]);
    let handled = false;
    const receipt = gate.waitForHandled("last-output").then((value) => {
      handled = value;
    });
    await gate.beforeEmit("last-output");
    expect(handled).toBe(false);
    gate.recordHandled("unrelated");
    await Promise.resolve();
    expect(handled).toBe(false);
    gate.recordHandled("last-output");
    await receipt;
    expect(handled).toBe(true);
    expect(await gate.waitForHandled("unconfigured")).toBe(false);
  });

  it("rejects duplicate labels", () => {
    expect(() => makeProviderReplayGate([], ["last-output", "last-output"])).toThrow(
      "Duplicate provider replay handled label last-output.",
    );
    expect(() => makeProviderReplayGate(["approval", "approval"])).toThrow(
      "Duplicate provider replay gate label approval.",
    );
  });
});
