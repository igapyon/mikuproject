import path from "node:path";

import { checkCliV1SchemaValidators } from "./lib/v1/cli-v1-schema-validator-generator.mjs";

if (process.argv.length !== 2) {
  throw new Error("usage: node scripts/check-cli-v1-schema-validators.mjs");
}

const { outputPath } = await checkCliV1SchemaValidators();
console.log(`[check:cli-v1-schema-validators] up to date ${path.relative(process.cwd(), outputPath)}`);
