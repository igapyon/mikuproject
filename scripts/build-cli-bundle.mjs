import fs from "node:fs";
import path from "node:path";
import { createRequire } from "node:module";

import { CORE_API_MODULE_RELATIVE_PATHS } from "./lib/core-api-loader.mjs";

const ROOT = process.cwd();
const requireFromRoot = createRequire(path.resolve(ROOT, "package.json"));

const args = parseArgs(process.argv.slice(2));
const outDir = path.resolve(
  args.out || path.join(ROOT, "bundle", "mikuproject-cli-bundle")
);
const bundleName = path.basename(outDir);

const rootPackageJson = readJson(path.resolve(ROOT, "package.json"));
const jsdomPackageJsonPath = resolveInstalledPackageJson("jsdom", requireFromRoot);
const jsdomPackageJson = readJson(jsdomPackageJsonPath);
const xmldomPackageJsonPath = resolveInstalledPackageJson("@xmldom/xmldom", requireFromRoot);
const xmldomPackageJson = readJson(xmldomPackageJsonPath);
const CLI_RUNTIME_DEPENDENCIES = ["jsdom", "@xmldom/xmldom"];

buildCliBundle();

function buildCliBundle() {
  fs.rmSync(outDir, { recursive: true, force: true });
  fs.mkdirSync(outDir, { recursive: true });

  for (const relativePath of getRuntimeFileRelativePaths()) {
    copyRepoFile(relativePath);
  }

  writeBundlePackageJson();
  writeBundleReadme();
  copyRuntimeNodeModules();

  console.log(`[build:cli-bundle] generated ${path.relative(ROOT, outDir)}`);
}

function parseArgs(argv) {
  const options = {};
  for (let index = 0; index < argv.length; index += 1) {
    const token = argv[index];
    if (!token.startsWith("--")) {
      throw new Error(`未対応の引数です: ${token}`);
    }

    const key = token.slice(2);
    const value = argv[index + 1];
    if (value === undefined || value.startsWith("--")) {
      throw new Error(`オプション ${token} には値が必要です`);
    }
    options[key] = value;
    index += 1;
  }
  return options;
}

function getRuntimeFileRelativePaths() {
  const files = new Set([
    "scripts/mikuproject-cli.mjs",
    "scripts/lib/core-api-loader.mjs",
    ...CORE_API_MODULE_RELATIVE_PATHS
  ]);
  return Array.from(files).sort();
}

function copyRepoFile(relativePath) {
  const sourcePath = path.resolve(ROOT, relativePath);
  const destinationPath = path.resolve(outDir, relativePath);
  fs.mkdirSync(path.dirname(destinationPath), { recursive: true });
  fs.copyFileSync(sourcePath, destinationPath);
}

function writeBundlePackageJson() {
  const document = {
    name: `${rootPackageJson.name}-cli-bundle`,
    version: rootPackageJson.version,
    private: true,
    description: "Self-contained mikuproject CLI bundle.",
    license: rootPackageJson.license,
    type: "module",
    bin: {
      mikuproject: "scripts/mikuproject-cli.mjs"
    },
    engines: {
      node: ">=18"
    },
    dependencies: {
      jsdom: jsdomPackageJson.version,
      "@xmldom/xmldom": xmldomPackageJson.version
    }
  };

  const packageJsonPath = path.resolve(outDir, "package.json");
  fs.writeFileSync(packageJsonPath, `${JSON.stringify(document, null, 2)}\n`, "utf8");
}

function writeBundleReadme() {
  const readmePath = path.resolve(outDir, "README.md");
  const lines = [
    `# ${bundleName}`,
    "",
    "この bundle は `mikuproject` CLI を追加の `npm install` なしで実行するための配布物です。",
    "",
    "## 使い方",
    "",
    "```bash",
    "node scripts/mikuproject-cli.mjs ai spec",
    "```",
    "",
    "または `npm link` / `npm install -g <展開先>` 相当で `mikuproject` を PATH に出せます。"
  ];
  fs.writeFileSync(readmePath, `${lines.join("\n")}\n`, "utf8");
}

function copyRuntimeNodeModules() {
  const sourceNodeModulesDir = path.resolve(ROOT, "node_modules");
  const destinationNodeModulesDir = path.resolve(outDir, "node_modules");
  fs.mkdirSync(destinationNodeModulesDir, { recursive: true });

  const visitedPackageJsonPaths = new Set();
  const queue = CLI_RUNTIME_DEPENDENCIES.map((name) => ({ name, requireFn: requireFromRoot }));

  while (queue.length > 0) {
    const current = queue.shift();
    const packageJsonPath = resolveInstalledPackageJson(current.name, current.requireFn);
    if (visitedPackageJsonPaths.has(packageJsonPath)) {
      continue;
    }
    visitedPackageJsonPaths.add(packageJsonPath);

    const packageJson = readJson(packageJsonPath);
    const packageDir = path.dirname(packageJsonPath);
    const relativePackageDir = path.relative(sourceNodeModulesDir, packageDir);
    if (relativePackageDir.startsWith("..")) {
      throw new Error(`node_modules 配下でない package を検出しました: ${packageDir}`);
    }

    const destinationPackageDir = path.resolve(destinationNodeModulesDir, relativePackageDir);
    fs.mkdirSync(path.dirname(destinationPackageDir), { recursive: true });
    fs.cpSync(packageDir, destinationPackageDir, {
      recursive: true,
      force: true,
      dereference: true
    });

    const packageRequire = createRequire(packageJsonPath);
    for (const dependencyName of Object.keys(packageJson.dependencies || {})) {
      queue.push({ name: dependencyName, requireFn: packageRequire });
    }
    for (const dependencyName of Object.keys(packageJson.optionalDependencies || {})) {
      if (!isPackageInstalled(dependencyName, packageRequire)) {
        continue;
      }
      queue.push({ name: dependencyName, requireFn: packageRequire });
    }
  }
}

function resolveInstalledPackageJson(packageName, requireFn) {
  try {
    return requireFn.resolve(`${packageName}/package.json`);
  } catch (_error) {
    try {
      const resolvedEntryPath = requireFn.resolve(packageName);
      return findNearestPackageJson(resolvedEntryPath, packageName);
    } catch (_innerError) {
      throw new Error(
        `${packageName} が見つかりません。先に repo root で npm install を実行してください`
      );
    }
  }
}

function isPackageInstalled(packageName, requireFn) {
  try {
    requireFn.resolve(`${packageName}/package.json`);
    return true;
  } catch (_error) {
    return false;
  }
}

function readJson(filePath) {
  return JSON.parse(fs.readFileSync(filePath, "utf8"));
}

function findNearestPackageJson(startPath, expectedPackageName) {
  let currentDir = path.dirname(startPath);
  while (true) {
    const packageJsonPath = path.join(currentDir, "package.json");
    if (fs.existsSync(packageJsonPath)) {
      const document = readJson(packageJsonPath);
      if (document.name === expectedPackageName) {
        return packageJsonPath;
      }
    }

    const parentDir = path.dirname(currentDir);
    if (parentDir === currentDir) {
      throw new Error(`package.json が見つかりません: ${expectedPackageName}`);
    }
    currentDir = parentDir;
  }
}
