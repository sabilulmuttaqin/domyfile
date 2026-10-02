import { defineConfig } from "vitest/config";

export default defineConfig({
  test: {
    // Playwright e2e specs live under tests/e2e/ and must never be collected
    // or transformed by vitest (they crash with "Playwright Test did not
    // expect test() to be called here" and fail the whole run).
    include: ["src/**/*.{test,spec}.?(c|m)[jt]s?(x)", "tests/**/*.{test,spec}.?(c|m)[jt]s?(x)"],
    exclude: [".build/**", "tests/e2e/**"],
  },
});
