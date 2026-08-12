import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

import { sha256RawBytes } from "../scripts/lib/v1/cli-v1-canonical-json.mjs";
import { isCliV1Error } from "../scripts/lib/v1/cli-v1-errors.mjs";
import { createV1DiagnosticFromError } from "../scripts/lib/v1/cli-v1-result.mjs";
import {
  semanticIssuesToV1Errors,
  validateV1SemanticState
} from "../scripts/lib/v1/cli-v1-semantic-validator.mjs";
import {
  MS_PROJECT_XML_ADAPTER,
  MS_PROJECT_XML_SUBSET_PROFILE,
  decodeMsProjectXmlSubset
} from "../scripts/lib/v1/cli-v1-xml-adapter.mjs";
import {
  validateArtifact,
  validateCliDiagnostic
} from "../scripts/generated/cli-v1-schema-validators.mjs";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const projectFixtureRoot = path.join(repoRoot, "testdata/conformance/v1/fixtures/project");
const goldenSemanticState = readJson("testdata/conformance/v1/golden/semantic/dependency.state.json");

describe("v1 MS Project XML subset adapter", () => {
  it("decodes canonical S-V001 XML into the exact semantic golden without mutating raw provenance", () => {
    const rawXml = readFixture("dependency-canonical.xml");
    const decoded = decodeMsProjectXmlSubset(rawXml);

    expect(decoded.format_profile).toBe(MS_PROJECT_XML_SUBSET_PROFILE);
    expect(decoded.adapter).toBe(MS_PROJECT_XML_ADAPTER);
    expect(decoded.raw_digest).toEqual(sha256RawBytes(rawXml));
    expect(decoded.normalizations).toEqual([]);
    expect(decoded.adapter_issues).toEqual([]);
    expect(decoded.state).toEqual(goldenSemanticState);
    expect(validateArtifact(decoded.state)).toBe(true);
    expect(validateV1SemanticState(decoded.state, { adapterIssues: decoded.adapter_issues })).toEqual({
      status: "valid",
      valid: true,
      issues: []
    });
  });

  it("keeps S-I012 as semantic.invalid with its stable rule and semantic location", () => {
    const decoded = decodeMsProjectXmlSubset(readFixture("dependency-percent-101.xml"));
    const validation = validateV1SemanticState(decoded.state, { adapterIssues: decoded.adapter_issues });

    expect(validation).toMatchObject({
      status: "invalid",
      valid: false,
      issues: [{
        code: "semantic.invalid",
        rule_id: "S-I012",
        path: "tasks[uid=2].percent_complete"
      }]
    });
    const [error] = semanticIssuesToV1Errors(validation);
    expect(isCliV1Error(error)).toBe(true);
    const diagnostic = createV1DiagnosticFromError(error);
    expect(diagnostic).toMatchObject({
      code: "semantic.invalid",
      category: "semantic",
      location: { scope: "semantic", path: "tasks[uid=2].percent_complete", rule_id: "S-I012" }
    });
    expect(validateCliDiagnostic(diagnostic)).toBe(true);
  });

  it("keeps S-I020 unsupported external data distinct from invalid state", () => {
    const decoded = decodeMsProjectXmlSubset(readFixture("dependency-unsupported-actual.xml"));
    const validation = validateV1SemanticState(decoded.state, { adapterIssues: decoded.adapter_issues });

    expect(validation).toMatchObject({
      status: "unsupported",
      valid: false,
      issues: [{
        code: "semantic.unsupported",
        rule_id: "S-I020",
        path: "tasks[uid=2].actual_start"
      }]
    });
  });

  it("records only a leading UTF-8 BOM normalization and rejects invalid text/encoding/profile inputs", () => {
    const canonicalText = readFixture("dependency-canonical.xml").toString("utf8");
    const bomDecoded = decodeMsProjectXmlSubset(Buffer.concat([Buffer.from([0xef, 0xbb, 0xbf]), Buffer.from(canonicalText)]));
    expect(bomDecoded.normalizations).toEqual([{
      code: "text.utf8-bom-removed",
      path: "Project",
      before: "utf8-bom",
      after: "no-bom"
    }]);
    expect(bomDecoded.state).toEqual(goldenSemanticState);

    expectRejected(Buffer.from([0xc3, 0x28]), "text.invalid-utf8");
    expectRejected(
      Buffer.from(canonicalText.replace('encoding="UTF-8"', 'encoding="UTF-16"')),
      "xml.encoding-unsupported"
    );
    expectRejected(
      Buffer.from(canonicalText.replace('xmlns="http://schemas.microsoft.com/project"', 'xmlns="urn:not-microsoft-project"')),
      "xml.profile-unsupported"
    );
  });

  it("fails closed for duplicate structural fields, unknown data, and semantic forest/dependency invariants", () => {
    const canonicalText = readFixture("dependency-canonical.xml").toString("utf8");
    expectRejected(
      Buffer.from(canonicalText.replace("<Name>Dependency Project</Name>", "<Name>Dependency Project</Name><Name>Duplicate</Name>")),
      "xml.invalid"
    );

    const unknown = decodeMsProjectXmlSubset(Buffer.from(canonicalText.replace("<PercentComplete>0</PercentComplete>", "<PercentComplete>0</PercentComplete><BaselineCost>1</BaselineCost>")));
    expect(validateV1SemanticState(unknown.state, { adapterIssues: unknown.adapter_issues })).toMatchObject({
      status: "unsupported",
      issues: [{ code: "semantic.unsupported", rule_id: "S-I020", path: "tasks[uid=2].baseline_cost" }]
    });

    const invalidState = structuredClone(goldenSemanticState);
    invalidState.tasks[0].summary = true;
    invalidState.dependencies.push({ predecessor_uid: "2", successor_uid: "1", type: "FS", lag: "PT0H0M0S" });
    const validation = validateV1SemanticState(invalidState);
    expect(validation.issues).toEqual(expect.arrayContaining([
      expect.objectContaining({ code: "semantic.invalid", rule_id: "S-I004" }),
      expect.objectContaining({ code: "semantic.invalid", rule_id: "S-I016" })
    ]));

    const invalidOutline = decodeMsProjectXmlSubset(Buffer.from(canonicalText.replace("<OutlineNumber>2</OutlineNumber>", "<OutlineNumber>9</OutlineNumber>")));
    expect(validateV1SemanticState(invalidOutline.state, { adapterIssues: invalidOutline.adapter_issues })).toMatchObject({
      status: "invalid",
      issues: [{ code: "semantic.invalid", rule_id: "S-I003", path: "tasks[uid=2].outline_number" }]
    });
  });
});

function readFixture(name) {
  return fs.readFileSync(path.join(projectFixtureRoot, name));
}

function readJson(relativePath) {
  return JSON.parse(fs.readFileSync(path.join(repoRoot, relativePath), "utf8"));
}

function expectRejected(rawXml, code) {
  try {
    decodeMsProjectXmlSubset(rawXml);
  } catch (error) {
    expect(isCliV1Error(error)).toBe(true);
    expect(error.code).toBe(code);
    return;
  }
  throw new Error(`expected ${code}`);
}
