import path from "node:path";

import { writeCliV1SchemaValidators } from "./lib/v1/cli-v1-schema-validator-generator.mjs";

if (process.argv.length !== 2) {
  throw new Error("usage: node scripts/generate-cli-v1-schema-validators.mjs");
}

const { outputPath } = await writeCliV1SchemaValidators();
console.log(`[generate:cli-v1-schema-validators] generated ${path.relative(process.cwd(), outputPath)}`);
