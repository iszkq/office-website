import { MetadataRoute } from "next";

export const dynamic = "force-static";

export default function robots(): MetadataRoute.Robots {
  return {
    rules: [
      {
        userAgent: "*",
        allow: "/",
        disallow: ["/settings", "/api/"],
      },
    ],
    sitemap: "https://124.222.193.241:6258/sitemap.xml",
  };
}
