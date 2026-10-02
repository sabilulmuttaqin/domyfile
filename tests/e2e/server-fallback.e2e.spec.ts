import { join } from "node:path";
import { test, expect, type Page } from "@playwright/test";

const fixture = (relativePath: string) => join(process.cwd(), "tests", "fixtures", relativePath);

const mockServerFallback = async (page: Page, params: {
  toolId: "PDF-01" | "VID-01";
  inputBytes: number;
  outputBytes: number;
  savingsPercent: number;
  outputMime: "application/pdf" | "video/webm";
  outputFormat: "PDF" | "WebM";
  resultDetails: Record<string, number>;
  output: number[];
}) => {
  const jobId = "b".repeat(32);
  const expiresAt = "2026-08-18T00:15:00.000Z";

  await page.route("**/v1/jobs**", async (route) => {
    const request = route.request();
    const url = new URL(request.url());

    if (request.method() === "POST" && url.pathname === "/__server-fallback/v1/jobs") {
      const body = request.postDataJSON() as Record<string, unknown>;
      expect(body).not.toHaveProperty("filename");
      expect(body).toMatchObject({ toolId: params.toolId, inputBytes: params.inputBytes });
      await route.fulfill({
        status: 201,
        contentType: "application/json",
        body: JSON.stringify({
          jobId,
          accessToken: "temporary-browser-test-token",
          uploadUrl: `${url.origin}/__server-fallback/v1/jobs/${jobId}/input`,
          expiresAt,
        }),
      });
      return;
    }

    if (request.method() === "PUT" && url.pathname.endsWith(`/v1/jobs/${jobId}/input`)) {
      await route.fulfill({ status: 202 });
      return;
    }

    if (request.method() === "GET" && url.pathname.endsWith(`/v1/jobs/${jobId}/download`)) {
      await route.fulfill({
        status: 200,
        contentType: params.outputMime,
        body: Buffer.from(params.output),
      });
      return;
    }

    if (request.method() === "GET" && url.pathname.endsWith(`/v1/jobs/${jobId}`)) {
      await route.fulfill({
        status: 200,
        contentType: "application/json",
        body: JSON.stringify({
          jobId,
          status: "ready",
          expiresAt,
          result: {
            inputBytes: params.inputBytes,
            outputBytes: params.outputBytes,
            savingsPercent: params.savingsPercent,
            outputMime: params.outputMime,
            outputFormat: params.outputFormat,
            ...params.resultDetails,
          },
        }),
      });
      return;
    }

    await route.continue();
  });
};

test("Compress PDF completes through the temporary server boundary and reports real savings", async ({ page }) => {
  await mockServerFallback(page, {
    toolId: "PDF-01",
    inputBytes: 656029,
    outputBytes: 498545,
    savingsPercent: 24.0,
    outputMime: "application/pdf",
    outputFormat: "PDF",
    resultDetails: { pageCount: 8, textPagesPreserved: 8 },
    output: [37, 80, 68, 70, 45, 49, 46, 55],
  });

  await page.goto("/pdf/compress-pdf/");
  await page.evaluate(() => { window.fetch = window.fetch.bind(window); });
  await expect(page.locator("[data-server-fallback='true']")).toHaveCount(1);
  await expect(page.locator("[data-processing-note]")).toContainText("uploaded temporarily");
  await page.locator("[data-file-input]").setInputFiles(fixture("compress-mixed-content.pdf"));
  await expect(page.locator("[data-process-button]")).toBeEnabled();
  await page.locator("[data-process-button]").click();

  await expect(page.locator("[data-result-card]")).toBeVisible();
  await expect(page.locator("[data-result-summary]")).toContainText("24% size change");
  await expect(page.locator("[data-result-list]")).toContainText("8 pages");
  const [download] = await Promise.all([
    page.waitForEvent("download"),
    page.locator("[data-result-download]").click(),
  ]);
  expect(download.suggestedFilename()).toBe("compressed-document.pdf");
});

test("Compress Video completes through the temporary server boundary and preserves reported media metadata", async ({ page }) => {
  await mockServerFallback(page, {
    toolId: "VID-01",
    inputBytes: 14820,
    outputBytes: 10838,
    savingsPercent: 26.87,
    outputMime: "video/webm",
    outputFormat: "WebM",
    resultDetails: { durationSeconds: 1.029, width: 160, height: 90, audioStreams: 1 },
    output: [0x1a, 0x45, 0xdf, 0xa3],
  });

  await page.goto("/video/compress-video/");
  await page.evaluate(() => { window.fetch = window.fetch.bind(window); });
  await expect(page.locator("[data-server-fallback='true']")).toHaveCount(1);
  await page.locator("[data-file-input]").setInputFiles(fixture("video/mp4-h264-aac.mp4"));
  await expect(page.locator("[data-process-button]")).toBeEnabled();
  await page.locator("[data-process-button]").click();

  await expect(page.locator("[data-result-card]")).toBeVisible();
  await expect(page.locator("[data-result-summary]")).toContainText("27% size change");
  await expect(page.locator("[data-result-list]")).toContainText("160 × 90 px");
  await expect(page.locator("[data-result-list]")).toContainText("1.03 s");
  const [download] = await Promise.all([
    page.waitForEvent("download"),
    page.locator("[data-result-download]").click(),
  ]);
  expect(download.suggestedFilename()).toBe("compressed-video.webm");
});
