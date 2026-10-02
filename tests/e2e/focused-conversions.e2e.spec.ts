import { readFile } from "node:fs/promises";
import { join } from "node:path";
import { expect, test, type Page } from "@playwright/test";

const fixture = (relativePath: string) => join(process.cwd(), "tests", "fixtures", relativePath);

const focusedRoutes = [
  { route: "/image/jpg-to-png/", title: "Convert JPG to PNG - Preserve Image Dimensions | DoMyFile", input: "JPG", output: "PNG", action: "Convert to PNG", related: "/image/jpg-to-webp/", h1: "Convert JPG or JPEG to PNG" },
  { route: "/image/jpg-to-webp/", title: "Convert JPG to WebP for Websites | DoMyFile", input: "JPG", output: "WebP", action: "Convert to WebP", related: "/image/jpg-to-png/", h1: "Convert JPG or JPEG to WebP" },
  { route: "/image/png-to-jpg/", title: "Convert PNG to JPG with Background Control | DoMyFile", input: "PNG", output: "JPG", action: "Convert to JPG", related: "/image/png-to-webp/", h1: "Convert PNG images to JPG" },
  { route: "/image/png-to-webp/", title: "Convert PNG to WebP with Quality Control | DoMyFile", input: "PNG", output: "WebP", action: "Convert to WebP", related: "/image/png-to-jpg/", h1: "Convert PNG images to WebP" },
  { route: "/image/webp-to-jpg/", title: "Convert WebP to JPG with Background Control | DoMyFile", input: "WebP", output: "JPG", action: "Convert to JPG", related: "/image/webp-to-png/", h1: "Convert WebP images to JPG" },
  { route: "/image/webp-to-png/", title: "Convert WebP to PNG without Resizing | DoMyFile", input: "WebP", output: "PNG", action: "Convert to PNG", related: "/image/webp-to-jpg/", h1: "Convert WebP images to PNG" },
  { route: "/pdf/jpg-to-pdf/", title: "Convert JPG Images to PDF | DoMyFile", input: "JPG", output: "PDF", action: "Convert to PDF", related: "/pdf/png-to-pdf/", h1: "Convert JPG or JPEG images to PDF" },
  { route: "/pdf/png-to-pdf/", title: "Convert PNG Images to PDF | DoMyFile", input: "PNG", output: "PDF", action: "Convert to PDF", related: "/pdf/jpg-to-pdf/", h1: "Convert PNG images to PDF" },
  { route: "/pdf/webp-to-pdf/", title: "Convert WebP Images to PDF | DoMyFile", input: "WebP", output: "PDF", action: "Convert to PDF", related: "/pdf/jpg-to-pdf/", h1: "Convert WebP images to PDF" },
] as const;

const readSignature = (bytes: Buffer) => {
  if (bytes.subarray(0, 5).toString("ascii") === "%PDF-") return "PDF";
  if (bytes[0] === 0xff && bytes[1] === 0xd8 && bytes[2] === 0xff) return "JPG";
  if (bytes.subarray(0, 8).compare(Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])) === 0) return "PNG";
  if (bytes.subarray(0, 4).toString("ascii") === "RIFF" && bytes.subarray(8, 12).toString("ascii") === "WEBP") return "WebP";
  return "unknown";
};

const downloadResult = async (page: Page) => {
  const [download] = await Promise.all([
    page.waitForEvent("download", { timeout: 120_000 }),
    page.locator("[data-download-all]").click(),
  ]);
  const outputPath = join(process.cwd(), "test-results", `focused-${Date.now()}-${download.suggestedFilename()}`);
  await download.saveAs(outputPath);
  return { path: outputPath, bytes: await readFile(outputPath), name: download.suggestedFilename() };
};

test("focused conversion routes expose unique intent metadata and fixed outputs", async ({ page }) => {
  for (const tool of focusedRoutes) {
    await page.goto(tool.route);
    await expect(page).toHaveTitle(tool.title);
    await expect(page.locator(".workspace-heading h1")).toContainText(tool.h1);
    const metaDescription = await page.locator("meta[name='description']").getAttribute("content");
    expect(metaDescription).toContain(tool.input);
    expect(metaDescription).toContain(tool.output);
    await expect(page.locator("link[rel='canonical']")).toHaveAttribute("href", tool.route);
    await expect(page.locator(".supported-formats")).toContainText(tool.input);
    await expect(page.locator("[data-process-button]")).toHaveText(tool.action);
    await expect(page.locator(".tool-info-card").first()).toContainText(`${tool.input} → ${tool.output}`);
    await expect(page.locator(".faq-list")).toContainText(tool.output);
    await expect(page.locator(`.related-tool-grid a[href='${tool.related}']`)).toHaveCount(1);
    await expect(page.locator("#outputFormat, #cropOutputFormat")).toHaveCount(0);
  }

  await page.goto("/tools/");
  await expect(page.locator("[data-tool-slug='image-converter'], [data-tool-slug='images-to-pdf']")).toHaveCount(0);
});

test("focused image and image-to-PDF routes reuse validated local engines", async ({ page }) => {
  test.setTimeout(300_000);
  const requestBodies: Array<Buffer | undefined> = [];
  page.on("request", (request) => requestBodies.push(request.postDataBuffer() ?? undefined));

  const imageJobs = [
    { route: "/image/jpg-to-webp/", input: "compression-photo.jpg", output: "WebP" },
    { route: "/image/png-to-webp/", input: "one-pixel.png", output: "WebP" },
    { route: "/image/webp-to-jpg/", input: "one-pixel.webp", output: "JPG" },
    { route: "/image/webp-to-png/", input: "one-pixel.webp", output: "PNG" },
  ] as const;

  for (const job of imageJobs) {
    await page.goto(job.route);
    await page.locator("[data-file-input]").setInputFiles(fixture(job.input));
    await expect(page.locator("[data-process-button]")).toBeEnabled();
    await page.locator("[data-process-button]").click();
    await expect(page.locator("[data-result-card]")).toBeVisible({ timeout: 120_000 });
    const result = await downloadResult(page);
    expect(readSignature(result.bytes)).toBe(job.output);
    expect(result.name.toLowerCase()).toContain(job.output.toLowerCase() === "webp" ? ".webp" : job.output === "PNG" ? ".png" : ".jpg");
  }

  const pdfJobs = [
    { route: "/pdf/jpg-to-pdf/", input: "compression-photo.jpg" },
    { route: "/pdf/png-to-pdf/", input: "one-pixel.png" },
    { route: "/pdf/webp-to-pdf/", input: "one-pixel.webp" },
  ] as const;

  for (const job of pdfJobs) {
    await page.goto(job.route);
    await page.locator("[data-file-input]").setInputFiles(fixture(job.input));
    await expect(page.locator("[data-process-button]")).toBeEnabled();
    await page.locator("[data-process-button]").click();
    await expect(page.locator("[data-result-card]")).toBeVisible({ timeout: 120_000 });
    const result = await downloadResult(page);
    expect(readSignature(result.bytes)).toBe("PDF");
    expect(result.name.toLowerCase()).toContain(".pdf");
  }

  expect(requestBodies.filter(Boolean)).toHaveLength(0);
});

test("focused routes reject the wrong input and old generic URLs redirect permanently", async ({ page, request }) => {
  await page.goto("/image/jpg-to-webp/");
  await page.locator("[data-file-input]").setInputFiles(fixture("one-pixel.png"));
  await expect(page.locator("[data-file-error]")).toContainText("not supported");
  await expect(page.locator("[data-process-button]")).toBeDisabled();

  for (const [oldRoute, newRoute] of [["/image/image-converter/", "/image/"], ["/pdf/images-to-pdf/", "/pdf/"]] as const) {
    const response = await request.get(oldRoute, { maxRedirects: 0 });
    expect([301, 308]).toContain(response.status());
    expect(response.headers().location).toContain(newRoute);
  }

  const sitemap = await (await request.get("/sitemap.xml")).text();
  expect(sitemap).toContain("/image/jpg-to-webp/");
  expect(sitemap).toContain("/pdf/webp-to-pdf/");
  expect(sitemap).not.toContain("/image/image-converter/");
  expect(sitemap).not.toContain("/pdf/images-to-pdf/");
});
