import { defineConfig } from "astro/config";
import { SITE_ORIGIN } from "./src/config/site";

export default defineConfig({
  output: "static",
  trailingSlash: "always",
  site: SITE_ORIGIN,
  redirects: {
    "/image/image-converter": `${SITE_ORIGIN}/image/`,
    "/pdf/images-to-pdf": `${SITE_ORIGIN}/pdf/`,
    // The remove-background tool is deferred and has no public route; keep old
    // links alive by redirecting the non-trailing-slash variant to /image/.
    "/image/remove-background": `${SITE_ORIGIN}/image/`,
  },
});
