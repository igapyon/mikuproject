# Miku Software MCP Design v20260426

This memo organizes design characteristics commonly expected for MCP server versions in the `miku` software series.

The initial versions of the related tools were created by `Mikuku` and Toshiki Iga.

The current contents are based on the miku main-application design memo, Java application design memo, straight-conversion guide, Agent Skills design memo, and the MCP concepts checked on 2026-04-26.

## Design Summary

The MCP server versions in the `miku` software series can be summarized as follows.

> Protocol adapters that expose miku product operations to MCP clients as tools, resources, and prompts, while preserving upstream semantics, local-first workflows, structured artifacts, diagnostics, and runtime traceability.

What characterizes MCP server versions is not a new product separate from the upstream application. It is a repeated set of constraints.

- Keep the semantic center of the upstream main application
- Expose product operations through stable MCP tools
- Expose reusable state, specifications, reports, and generated files through MCP resources
- Expose a small number of product-specific prompts when useful
- Keep local stdio execution as the first implementation target
- Allow HTTP server deployment as a later transport and hosting variant
- Use upstream public APIs, CLI, or bundled runtime artifacts rather than duplicating product logic
- Treat the MCP server as a thin protocol adapter
- Keep local and server deployments aligned around the same tool contracts
- Preserve diagnostics, warnings, and artifact roles in structured results

## Role of This Document

This document is not a detailed specification for one MCP repository. Repository-specific behavior should be checked in each MCP server README, package metadata, tool schema, and product-specific documents.

Use the shared design documents together as follows.

- `docs/miku-soft-10-mainapp-design-v20260425.md`
  - describes the upstream product design and semantic center
- `docs/miku-soft-20-javaapp-design-v20260425.md`
  - describes Java runtime versions when they exist
- `docs/miku-soft-30-straight-conversion-v20260425.md`
  - describes how Java versions are created from upstream main applications
- `docs/miku-soft-40-agentskills-design-v20260425.md`
  - describes how Agent Skills versions expose miku workflows to AI agents
- `docs/miku-soft-50-mcp-design-v20260426.md`
  - describes how MCP server versions should expose miku workflows to MCP clients

This document separates the following levels.

- **Cross-cutting principles**: design policies that should generally be kept for MCP server versions.
- **Recommended conventions**: transport, tool, resource, state, runtime, and packaging shapes that make maintenance easier.
- **Observed tendencies**: design habits visible in current miku repositories and MCP usage.
- **Product-specific notes**: design decisions for specific MCP products, recorded as concrete examples.

When a specification is unclear, first check whether the decision preserves upstream meaning and keeps the MCP surface stable for clients. MCP convenience is important, but it should not hide unsupported behavior, discard diagnostics, or make local and server deployments drift into different products.

## Scope

The series whose names start with `miku` includes several related product types.

Main applications:

- `miku-abc-player`
- `miku-docx2md`
- `miku-indexgen`
- `miku-unicode-guard`
- `miku-xlsx2md`
- `mikuproject`
- `mikuscore`

Java application versions:

- `miku-indexgen-java`
- `mikuproject-java`
- `miku-xlsx2md-java`

Agent Skills versions:

- `mikuproject-skills`
- `mikuscore-skills`

MCP server versions:

- `mikuproject-mcp`

Projects with the `-mcp` suffix are positioned as MCP server adapters for original suffixless tools or their runtime artifacts.

This document focuses on MCP server versions. It does not define the Web UI conventions for upstream main applications, Java packaging conventions for Java application versions, or skill packaging conventions for Agent Skills repositories.

## Shared Direction

MCP server versions continue the shared direction of the miku series: small local bridge tools that convert, extract, inspect, normalize, or export domain files and domain data.

The MCP version emphasizes the parts that fit MCP operation particularly well.

- model-callable tools
- application-readable resources
- user-selectable prompts where useful
- structured JSON input and output schemas
- local import / validate / export operations
- resource links for generated artifacts
- diagnostics and warnings as structured results
- repeatable file-based operation
- stdio local server execution
- optional HTTP server deployment when remote access is required

An MCP version should not become a generic planner, generic converter, hosted replacement application, or independent semantic implementation. Its value is that MCP clients can use the upstream miku product correctly through stable tool and resource contracts.

## Relationship to Main Applications

MCP server versions are downstream of miku main applications.

The upstream main application owns the product semantics, canonical source, core conversions, output policy, and limitations. The MCP server owns protocol exposure, tool schema, resource naming, transport handling, runtime discovery, and result formatting.

The preferred relationship is as follows.

- upstream main application provides core APIs, CLI, bundled runtime, or stable artifacts
- MCP server calls those upstream surfaces
- MCP server exposes product operations as tools
- MCP server exposes reusable state and artifacts as resources
- MCP server may expose product-specific workflow prompts
- MCP server returns diagnostics, warnings, hard errors, and generated artifact links in a structured form
- MCP server does not reimplement core conversion behavior unless no upstream surface exists

If an upstream capability is missing, the preferred order is to request or implement the capability on the upstream side, expose it as a stable upstream API, and then let the MCP server call it. MCP servers should not silently add upstream product capabilities on the MCP side. MCP-local workaround code should be treated as temporary or product-specific, not as the new semantic center.

## Relationship to Agent Skills

Agent Skills and MCP servers are closely related, but they are not the same layer.

Agent Skills are instruction and workflow packages for agents. They explain when to activate a product-specific workflow, which runtime artifacts to use, how to manage intermediate files, and how to avoid confusing artifact roles.

MCP servers are protocol adapters. They expose callable tools, readable resources, and optional prompts to MCP clients.

The preferred relationship is as follows.

- Agent Skills may continue to guide agents that can read `SKILL.md` and run local commands directly
- MCP servers should serve clients that prefer a standard tool/resource/prompt protocol
- both should call the same upstream runtime artifacts or public APIs when possible
- both should use the same product vocabulary and artifact roles
- neither should become the semantic owner of the product

For a product such as `mikuproject`, an Agent Skills repository may include a local MCP server, or a separate `mikuproject-mcp` repository may provide one. In either case, the MCP contract should stay aligned with the Agent Skills operation map and upstream CLI contracts.

## Role of MCP Servers

In this document, an MCP server means a `miku` repository or package with an MCP server entrypoint that provides product operations through MCP.

An MCP server usually exposes one or more of the following.

- stdio server entrypoint for local MCP clients
- HTTP server entrypoint for hosted or remote MCP clients
- tool definitions with JSON input schemas
- structured tool results and optional output schemas
- resource definitions for state, specifications, reports, and generated files
- resource templates for session or workspace-scoped artifacts
- prompt definitions for product-specific workflows
- runtime adapters for Java CLI, Node.js CLI, or public APIs
- smoke tests for protocol startup and core tool calls

The MCP server is not primarily the browser UI, not primarily a Java port, and not primarily an Agent Skill. It can use Node.js, Java, or another runtime if that is the natural way to call the upstream product, but the MCP layer itself should remain thin.

The center of an MCP server is reliable protocol adaptation: receive a typed tool call, validate arguments, call the upstream runtime, preserve useful state, expose resulting artifacts as resources, and return diagnostics without hiding important constraints.

## Cross-Cutting Principles

miku MCP servers emphasize the following cross-cutting principles.

1. Preserve the semantic center and product boundary of the upstream main application.
2. Treat MCP as a protocol adapter, not as a replacement implementation.
3. Keep MCP tool contracts stable across local stdio and HTTP server deployments.
4. Prefer local stdio execution as the first implementation target.
5. Add HTTP server deployment only with explicit session, storage, authentication, and isolation boundaries.
6. Use structured input schemas, structured results, diagnostics, and resource links instead of loose prose.
7. Distinguish canonical source, conversation state, MCP resources, primary output, reports, and temporary files.
8. Make runtime selection, artifact storage, smoke testing, and upstream synchronization explainable.

### Basic Philosophy of MCP Servers

MCP server versions emphasize the following philosophy.

- Run product operations locally when practical
- Keep the upstream product meaning recognizable
- Use upstream API or CLI before writing MCP-local conversion logic
- Expose only product-specific operations that are useful to MCP clients
- Keep tool names stable, explicit, and product-prefixed
- Keep intermediate JSON internal when direct tool composition is possible
- Return generated files as resources or resource links when practical
- Preserve diagnostics and warnings as part of the result
- Keep state and generated files easy to save, compare, and rerun
- Keep local and hosted deployments aligned around the same tool contracts

The value of an MCP server version is not that it can answer a domain question conversationally in a generic way. It is that MCP clients can safely call specific miku workflows through structured operations that are faithful to the upstream product.

### Common Principles for MCP Servers

MCP servers use the following principles as defaults.

- Use a `-mcp` suffix when the MCP server is a separate repository or package
- Keep the product name in tool names, such as `mikuproject.apply_patch`
- Use stdio transport as the first local implementation
- Treat HTTP transport as an additional deployment form, not a separate product contract
- Use Node.js or TypeScript when that is the simplest way to implement the MCP server
- Prefer bundled Java CLI and Node.js CLI runtime artifacts when local, reproducible execution needs them
- Treat the Java CLI runtime artifact as a single jar
- Treat the Node.js CLI runtime artifact as a single JavaScript file
- Prefer upstream public APIs, stable global APIs, or documented CLI commands
- Keep tool schemas and result schemas under version control
- Keep repository-level README user-facing and developer details under `docs/`
- Use `workplace/` or a configured workspace root for local scratch data, uploaded files, generated outputs, and verification files

These are defaults for the miku MCP series. Individual products may add product-specific conventions, but should not change these foundations casually.

## Transport Principles

MCP is a protocol, not a synonym for an HTTP server.

The first transport for miku MCP servers should usually be local stdio.

Local stdio is appropriate when:

- the user runs an MCP client on the same machine
- the product reads and writes local files
- the upstream runtime is available as a local jar, JavaScript file, or CLI command
- the workflow benefits from local-first processing and simple installation

HTTP transport is appropriate when:

- multiple clients need remote access
- the server owns session storage
- files are uploaded to a controlled workspace
- authentication, quota, and audit requirements are explicit
- the deployment environment can isolate user workspaces safely

Local stdio and HTTP deployments should share the same tool names and core input/output contracts. Differences should be limited to transport, file reference handling, session identity, storage, and authentication.

## Local Stdio Server Principles

The local stdio server is the preferred MVP shape.

It should:

- start as a local process launched by the MCP client
- communicate over standard input and standard output
- avoid requiring a network listener
- discover bundled or configured runtime artifacts
- call upstream CLI or APIs with explicit file paths
- confine file reads and writes to declared workspace roots where practical
- return generated file references as resources or resource links

The local stdio server should not behave like a long-running public web service. It is a local adapter process whose authority comes from the user and the MCP client configuration.

## HTTP Server Principles

The HTTP server is a deployment variant, not the initial semantic center.

An HTTP MCP server must define the following before it is treated as production-ready.

- authentication and authorization model
- session identity
- workspace isolation
- uploaded file lifecycle
- generated artifact lifecycle
- maximum input and output sizes
- cleanup policy
- audit or trace policy when needed
- runtime sandboxing or equivalent execution boundary
- user-visible confirmation policy for destructive operations when the client supports it

HTTP deployment should not expose arbitrary local filesystem paths from the host. It should prefer session-scoped resource URIs, upload identifiers, or controlled workspace-relative paths.

## Tool Design Principles

MCP tools are the main execution surface.

Tool names should be product-prefixed and stable.

Examples:

- `mikuproject.ai_spec`
- `mikuproject.detect_kind`
- `mikuproject.state_from_draft`
- `mikuproject.project_overview`
- `mikuproject.task_edit`
- `mikuproject.phase_detail`
- `mikuproject.validate_patch`
- `mikuproject.apply_patch`
- `mikuproject.state_diff`
- `mikuproject.export_workbook_json`
- `mikuproject.report_mermaid`

Tool input should use JSON Schema. Tool results should return structured content when practical. For compatibility with clients that primarily display text, structured results may also be serialized into a text content block.

Tools should prefer explicit argument names over positional behavior. Large input and output should be passed by file path, resource URI, upload ID, or resource link instead of being forced into chat-visible text.

## Resource Design Principles

MCP resources should expose reusable state and artifacts.

Typical resources include:

- AI-facing specifications
- current workbook state
- saved workbook states
- imported canonical source files
- generated report files
- diagnostics logs
- operation summaries

Resource URIs should distinguish artifact roles.

Examples:

- `mikuproject://spec/ai-json`
- `mikuproject://state/current`
- `mikuproject://state/{name}`
- `mikuproject://report/{name}`
- `mikuproject://diagnostics/{operationId}`

When a resource maps to a real local file, the server may return a `file://` resource link only when the client is expected to access that file directly. Otherwise, a product-specific URI should be preferred so that the server remains the controlled access point.

## Prompt Design Principles

MCP prompts should be small and product-specific.

Prompts are useful when a client wants user-selectable workflow templates. They should not become the primary location for product semantics.

Typical prompts include:

- create a new product-specific draft from user requirements
- revise an existing state using a small projection and patch
- review a product artifact using documented diagnostics

For products that already expose AI-facing specifications through upstream APIs or CLI commands, the MCP server should expose that specification as a tool or resource. Prompt text can reference the specification, but should not duplicate a large schema if the upstream product already provides a stable retrieval path.

## Runtime Principles

MCP servers should use declared runtime artifacts before broad repository exploration.

The preferred order is as follows.

1. Check the MCP server configuration
2. Check bundled runtime artifacts declared by the package
3. Use upstream public APIs, stable globals, or documented CLI/runtime flows
4. Use MCP-local helpers only when they are the intended adapter layer
5. Search broadly for alternatives only when the declared runtime path is missing or unusable

For products with both Java and Node.js runtime artifacts, the MCP server may prefer the Java CLI first for local execution. A single jar is easy to locate, smoke-test, and run in automation environments. If the Java CLI artifact is missing, cannot be invoked, or does not support the requested operation, the MCP server may fall back to the Node.js CLI runtime.

This runtime preference is an execution policy, not a semantic priority. The upstream main application remains the semantic center. Runtime differences should be reported as capability or compatibility diagnostics.

## State and Artifact Principles

MCP servers distinguish several kinds of data.

- canonical source or semantic base owned by the upstream product
- MCP server session state
- AI-facing projection or draft document
- patch or edit document returned by AI
- primary product output
- report or presentation output
- diagnostics and warnings
- temporary files and local scratch outputs

These roles should not be collapsed only because they are all represented as JSON or files.

For local stdio deployment, file paths can be a practical state boundary. For HTTP deployment, session-scoped resource URIs or upload identifiers are usually safer than arbitrary host paths.

For products such as `mikuproject`, `mikuproject_workbook_json` may be the practical MCP state handoff format, while `MS Project XML` remains the product's semantic base. The MCP server should make that distinction visible in naming, docs, and tool descriptions.

## Local and Server Variant Principles

A separate MCP product may support both local and server deployments.

The recommended split is:

- core tool definitions
- core input and result schemas
- runtime adapter interface
- local stdio entrypoint
- HTTP entrypoint
- storage adapter
- session adapter

Local and server variants should not create different operation vocabularies. They should share the same core tool names and result shapes.

Variant-specific differences should be limited to:

- transport
- file reference representation
- workspace and session scope
- authentication
- storage persistence
- artifact lifetime
- runtime isolation

If a capability is not safe or practical in one deployment form, the server should report an explicit capability or policy error rather than silently changing behavior.

## Error Handling Principles

MCP servers should separate the following.

- invalid tool arguments
- unsupported document kind
- missing base state
- upstream runtime failure
- upstream validation error
- upstream warning
- file access or storage policy error
- transport or session error

Hard errors should make the tool call fail in a way the MCP client can surface clearly.

Soft errors and upstream warnings should be returned in structured results when the operation can still produce a meaningful artifact. They should not be discarded only because the requested output file was generated.

## Security and Trust Principles

MCP servers expose executable operations to AI clients, so security boundaries must be explicit.

Local stdio servers should:

- avoid reading arbitrary files unless the user or client provides them explicitly
- avoid writing outside configured workspace or output roots when practical
- avoid accepting shell fragments or raw command lines from tool arguments
- call upstream runtimes through fixed argument arrays, not string-built shell commands
- report generated paths clearly

HTTP servers should additionally:

- authenticate clients or users
- isolate sessions
- restrict file access to controlled storage
- enforce upload and output size limits
- clean up temporary artifacts
- prevent cross-user resource access
- treat uploaded files as untrusted input

MCP tool descriptions and annotations are not a substitute for server-side validation.

## Product Boundary Principles

MCP servers must preserve the product boundary of the upstream miku tool.

Examples:

- `mikuproject-mcp` exposes WBS import, draft, patch, state, and report operations; it is not a full project-management server.
- `mikuscore-mcp` would expose score conversion, inspection, and handoff operations; it would not become a full notation editor.
- `miku-xlsx2md-mcp` would expose workbook-to-Markdown extraction and diagnostics; it would not become a general spreadsheet automation server.

If the MCP server starts needing broad domain features that the upstream application does not own, treat that as a product design issue rather than hiding it in MCP tool code.

## Recommended MVP Shape

For a first miku MCP server, use this shape.

- local stdio transport
- product-prefixed tools
- JSON Schema input definitions
- structured tool results
- resource links for generated artifacts
- bundled or configured runtime artifact lookup
- smoke test for server startup
- smoke test for one read-only tool
- smoke test for one state-changing tool with temporary files

For `mikuproject-mcp`, the MVP tools should be close to the Agent Skills MVP operation set.

- `mikuproject.ai_spec`
- `mikuproject.detect_kind`
- `mikuproject.state_from_draft`
- `mikuproject.project_overview`
- `mikuproject.task_edit`
- `mikuproject.phase_detail`
- `mikuproject.validate_patch`
- `mikuproject.apply_patch`
- `mikuproject.state_diff`
- `mikuproject.state_summarize`
- `mikuproject.export_workbook_json`

File import/export and report generation can be added after the core state workflow is stable.

## Out of Scope for the First Version

The first version should not include:

- broad hosted multi-tenant operation
- custom user management
- external AI model invocation
- API key management for model providers
- a browser UI replacement
- full domain editing outside upstream-supported operations
- automatic workflow planning beyond exposed tools and prompts
- broad filesystem browsing
- parallel implementation of upstream conversion logic

These may become useful later, but they should not be allowed to blur the initial MCP contract.

## Product-Specific Note: mikuproject-mcp

`mikuproject-mcp` should expose `mikuproject` operations to MCP clients.

The semantic center remains upstream `mikuproject`. The practical MCP state handoff format may be `mikuproject_workbook_json`. `MS Project XML` remains the semantic base and compatibility format for the product.

`mikuproject-mcp` may be developed as an independent repository and product, rather than as a subdirectory of `mikuproject-skills`.

This independence means that `mikuproject-mcp` owns its MCP-specific entrypoints, schemas, transport handling, and adapter code. It does not mean that `mikuproject-mcp` owns the product semantics of `mikuproject`.

The preferred repository relationship is as follows.

- `mikuproject`
  - owns the upstream product semantics
  - owns the canonical domain behavior, conversion policy, report policy, and AI-facing product contracts
- `mikuproject-java`
  - provides the Java runtime and jar-based execution path when used
  - keeps Java CLI behavior aligned with the upstream product
- `mikuproject-skills`
  - provides the Agent Skills workflow design reference
  - records operation maps, state-handling rules, file workflow rules, and agent-facing boundaries
- `mikuproject-mcp`
  - owns the MCP server entrypoints
  - owns tool, resource, and prompt definitions
  - owns MCP input and output schemas
  - owns local stdio transport and optional HTTP transport
  - owns session, workspace, storage, and runtime adapter policy for MCP deployments

For development, `mikuproject-mcp` may use local upstream checkouts under `workplace/upstream/`.

Recommended layout:

```text
mikuproject-mcp/
  workplace/
    upstream/
      mikuproject/
      mikuproject-java/
      mikuproject-skills/
```

These checkouts are reference material and local development inputs. They should normally remain outside Git tracking except for placeholder files needed to keep the workspace directory shape.

The upstream checkouts should be used to inspect and synchronize the following.

- upstream `mikuproject` CLI contracts and AI-facing specifications
- upstream `mikuproject` product behavior and artifact roles
- `mikuproject-java` jar behavior and Java CLI parity
- `mikuproject-skills` operation map and workflow boundary rules
- runtime artifact generation and smoke-test expectations

The MCP repository may copy, bundle, or download runtime artifacts according to its own packaging policy, but it should not treat `workplace/upstream/` itself as the installed runtime surface.

The clean responsibility split is:

```text
mikuproject
  -> upstream semantic owner

mikuproject-java
  -> Java runtime / jar provider

mikuproject-skills
  -> Agent Skills workflow design reference

mikuproject-mcp
  -> MCP protocol adapter
```

The first implementation should borrow heavily from `mikuproject-skills` design material, especially:

- operation names and operation grouping
- `mikuproject_workbook_json` as the practical state boundary
- Java runtime preference with Node.js runtime fallback
- distinction between structural workbook `XLSX` and `WBS XLSX` report output
- distinction between draft, patch, projection, workbook, report, and diagnostics artifacts
- default local output discipline for state, report, and temporary files

The first implementation should prefer local stdio transport. It should call `mikuproject.jar` or `mikuproject.mjs` through fixed runtime adapter code. It should keep the tool names aligned with the existing `mikuproject-skills` operation map and the Java / Node.js CLI correspondence.

If `mikuproject-mcp` later adds an HTTP server, it should keep the same core tool names. The HTTP server should add session-scoped resources, upload handling, authentication, storage policy, and artifact lifecycle management rather than changing the product operation vocabulary.
