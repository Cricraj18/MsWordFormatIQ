import { describe, expect, it } from "vitest";
import { APP_NAME, getTitleText } from "./version";

// Smoke test: proves Vitest and TypeScript are wired up correctly.
describe("setup smoke test", () => {
  it("has the right app name", () => {
    expect(APP_NAME).toBe("formatIQ");
  });

  it("builds the title text", () => {
    expect(getTitleText()).toBe("formatIQ v0.0.1");
  });
});
