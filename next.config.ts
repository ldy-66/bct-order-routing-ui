import type { NextConfig } from "next";

const isGitHubPages = process.env.GITHUB_PAGES === "true";
const githubPagesBasePath = "/bct-order-routing-ui";

const nextConfig: NextConfig = {
  ...(isGitHubPages ? {
    output: "export" as const,
    basePath: githubPagesBasePath,
    assetPrefix: githubPagesBasePath,
  } : {}),
};

export default nextConfig;
