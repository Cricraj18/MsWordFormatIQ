import { defineConfig } from "vitest/config";

// Vite builds the website. Vitest (tests) reads the "test" block below.
export default defineConfig({
  build: {
    target: "es2022",
  },
  worker: {
    format: "es",
  },
  test: {
    environment: "node",
    include: ["src/**/*.test.ts"],
  },
});
