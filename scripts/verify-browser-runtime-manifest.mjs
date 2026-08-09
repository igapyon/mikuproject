import { createHash } from "node:crypto";
import fs from "node:fs";
import path from "node:path";

const args = parseArgs(process.argv.slice(2));
const manifestPath = path.resolve(args.manifestPath);
const runtimePath = path.resolve(args.runtimePath);
const manifest = JSON.parse(fs.readFileSync(manifestPath, "utf8"));

if (manifest.schema_version !== "miku-project.browser-runtime-lock/v1") {
  throw new Error(`unsupported browser runtime lock schema: ${manifest.schema_version}`);
}
if (args.expectedReleaseTag && manifest.release_tag !== args.expectedReleaseTag) {
  throw new Error(`release tag mismatch: expected ${args.expectedReleaseTag}, actual ${manifest.release_tag}`);
}
if (path.basename(runtimePath) !== manifest.asset_name) {
  throw new Error(`runtime asset name mismatch: expected ${manifest.asset_name}, actual ${path.basename(runtimePath)}`);
}
if (!/^[0-9a-f]{64}$/.test(manifest.sha256)) {
  throw new Error("browser runtime lock sha256 must be 64 lowercase hexadecimal characters");
}

const actualSha256 = createHash("sha256").update(fs.readFileSync(runtimePath)).digest("hex");
if (actualSha256 !== manifest.sha256) {
  throw new Error(`runtime sha256 mismatch: expected ${manifest.sha256}, actual ${actualSha256}`);
}

console.log(`[verify:browser-runtime-manifest] ok ${manifest.release_tag} ${manifest.asset_name}`);

function parseArgs(argv) {
  if (argv.length < 2 || argv[0].startsWith("--") || argv[1].startsWith("--")) {
    throw new Error("usage: node scripts/verify-browser-runtime-manifest.mjs <manifest.json> <runtime.mjs> [--expected-release-tag <tag>]");
  }
  const options = {
    manifestPath: argv[0],
    runtimePath: argv[1]
  };
  for (let index = 2; index < argv.length; index += 1) {
    const token = argv[index];
    if (token !== "--expected-release-tag") {
      throw new Error(`unknown argument: ${token}`);
    }
    const value = argv[index + 1];
    if (value === undefined || value.startsWith("--")) {
      throw new Error(`${token} requires a value`);
    }
    options.expectedReleaseTag = value;
    index += 1;
  }
  return options;
}
