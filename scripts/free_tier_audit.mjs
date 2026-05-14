import fs from "node:fs";
import path from "node:path";

const ROOT = process.cwd();
const PUBLIC_DIR = path.join(ROOT, "public");

const CLOUDFLARE_PAGES_FREE_LIMITS = {
  maxFiles: 20_000,
  maxAssetBytes: 25 * 1024 * 1024,
};

const paidFeaturePaths = [
  "functions",
  "public/functions",
  "_worker.js",
  "public/_worker.js",
  "wrangler.toml",
];

function walkFiles(dir) {
  const files = [];
  const entries = fs.readdirSync(dir, { withFileTypes: true });

  for (const entry of entries) {
    const fullPath = path.join(dir, entry.name);
    if (entry.isDirectory()) {
      files.push(...walkFiles(fullPath));
    } else if (entry.isFile()) {
      files.push(fullPath);
    }
  }

  return files;
}

function existsRelative(relativePath) {
  return fs.existsSync(path.join(ROOT, relativePath));
}

const findings = [];

if (!fs.existsSync(PUBLIC_DIR)) {
  findings.push("Cloudflare Pages output directory public/ does not exist.");
} else {
  const files = walkFiles(PUBLIC_DIR);

  if (files.length > CLOUDFLARE_PAGES_FREE_LIMITS.maxFiles) {
    findings.push(
      `public/ has ${files.length.toLocaleString("ko-KR")} files, exceeding the Cloudflare Pages Free limit of ${CLOUDFLARE_PAGES_FREE_LIMITS.maxFiles.toLocaleString("ko-KR")}.`,
    );
  }

  for (const file of files) {
    const stat = fs.statSync(file);
    if (stat.size > CLOUDFLARE_PAGES_FREE_LIMITS.maxAssetBytes) {
      findings.push(
        `${path.relative(ROOT, file)} is ${(stat.size / 1024 / 1024).toFixed(1)} MiB, exceeding the Cloudflare Pages Free single-asset limit of 25 MiB.`,
      );
    }
  }
}

for (const relativePath of paidFeaturePaths) {
  if (existsRelative(relativePath)) {
    findings.push(
      `${relativePath} exists. This project should stay static-only to avoid Workers/Pages Functions usage and related quotas.`,
    );
  }
}

if (findings.length) {
  console.error("Free-tier audit failed:");
  for (const finding of findings) {
    console.error(`- ${finding}`);
  }
  process.exit(1);
}

console.log("Free-tier audit passed");
