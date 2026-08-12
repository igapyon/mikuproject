import {
  buildIoDiagnostics,
  describeBinaryInputForDiagnostics,
  ensureBinaryInputSource,
  ensureBinaryOutputTarget,
  ensureSingleStdinSource,
  readBinaryInput,
  readTextInput
} from "./cli-io.mjs";
import { buildCommandDiagnostics, parseDiagnosticsFormat } from "./cli-diagnostics.mjs";
import { ensureWorkbookJson, parsePlainJson } from "./cli-command-utils.mjs";

export async function runExchangeCommand(command, options, api) {
  const [scope, action, subject] = command;

  if (scope === "export" && subject === undefined) {
    ensureSingleStdinSource([
      { optionName: "--in", value: options.in, allowImplicitStdin: true }
    ]);
    const diagnosticsFormat = parseDiagnosticsFormat(options.diagnostics);
    if (action === "workbook-json") {
      const stateDocument = parsePlainJson(await readTextInput(options.in), "export workbook-json");
      ensureWorkbookJson(api, stateDocument, "export workbook-json");
      const model = api.workbookJson.importAsProjectModel(stateDocument).model;
      const exported = api.workbookJson.exportDocument(model);
      return {
        output: `${JSON.stringify(exported, null, 2)}\n`,
        diagnostics: diagnosticsFormat === "json"
          ? buildCommandDiagnostics("export workbook-json", {
            io: buildIoDiagnostics({
              inputs: [{ optionName: "--in", value: options.in, allowImplicitStdin: true }],
              output: options.out
            }),
            output_kind: "workbook_json",
            sheet_count: Object.keys(exported.sheets || {}).length
          })
          : []
      };
    }

    if (action === "xml") {
      const stateDocument = parsePlainJson(await readTextInput(options.in), "export xml");
      ensureWorkbookJson(api, stateDocument, "export xml");
      const model = api.workbookJson.importAsProjectModel(stateDocument).model;
      const output = `${api.msProject.exportToXml(model)}\n`;
      return {
        output,
        diagnostics: diagnosticsFormat === "json"
          ? buildCommandDiagnostics("export xml", {
            io: buildIoDiagnostics({
              inputs: [{ optionName: "--in", value: options.in, allowImplicitStdin: true }],
              output: options.out
            }),
            output_kind: "ms_project_xml",
            output_length: output.length
          })
          : []
      };
    }

    if (action === "xlsx") {
      ensureBinaryOutputTarget(options, "export xlsx");
      const stateDocument = parsePlainJson(await readTextInput(options.in), "export xlsx");
      ensureWorkbookJson(api, stateDocument, "export xlsx");
      const model = api.workbookJson.importAsProjectModel(stateDocument).model;
      const workbook = api.xlsx.exportWorkbook(model);
      const output = api.xlsx.encodeWorkbook(workbook);
      return {
        output,
        diagnostics: diagnosticsFormat === "json"
          ? buildCommandDiagnostics("export xlsx", {
            io: buildIoDiagnostics({
              inputs: [{ optionName: "--in", value: options.in, allowImplicitStdin: true }],
              output: options.out,
              outputBase64: options["out-base64"]
            }),
            output_kind: "xlsx",
            byte_length: output.length
          })
          : []
      };
    }
    return undefined;
  }

  if (scope === "import" && subject === undefined) {
    const diagnosticsFormat = parseDiagnosticsFormat(options.diagnostics);
    if (action === "xlsx") {
      ensureBinaryInputSource(options, "import xlsx");
      const sourceBytes = await readBinaryInput(options, "import xlsx");
      const imported = api.importExternal({
        source: { format: "xlsx", bytes: sourceBytes },
        mode: "replace"
      });
      const workbookDocument = api.workbookJson.exportDocument(imported.model);
      return {
        output: `${JSON.stringify(workbookDocument, null, 2)}\n`,
        diagnostics: diagnosticsFormat === "json"
          ? buildCommandDiagnostics("import xlsx", {
            io: buildIoDiagnostics({
              inputs: [describeBinaryInputForDiagnostics(options)],
              output: options.out
            }),
            input_kind: "xlsx",
            output_kind: "workbook_json",
            byte_length: sourceBytes.length,
            sheet_count: Object.keys(workbookDocument.sheets || {}).length
          })
          : []
      };
    }
  }

  return undefined;
}
