# Research MCP agent benchmark

Current as of 2026-08-25: the server exposes eleven tools
(`set_project_collection` was added after this file last said ten). The latest
full synthetic versioned run is `benchmark_runs/g16-full` (16/16
`passed_guarded`), recorded before the DOI/collection expansion, so it is
historical evidence for a smaller surface and not a current baseline.

This suite measures whether a local model can complete one paper
task at a time through the real MCP boundary, and how many server-reported tokens
it consumes.

The default run executes read-only tasks and reports the one still-planned
capability. It compares two conditions:

- `all`: every current MCP schema is visible to the agent;
- `minimal`: only the tool(s) required by that task are visible.

That comparison measures the token and tool-selection tradeoff of a wider MCP
surface. Each task is independently verified; the model's prose claim is never
treated as success.

## Commands

Read-only smoke benchmark:

```powershell
.\.venv\Scripts\python.exe -m research_mcp.benchmark
```

Current implemented surface, including disposable versioned Word writes:

```powershell
.\.venv\Scripts\python.exe -m research_mcp.benchmark `
  --tiers read_only versioned planned
```

Include the live Word/Zotero citation integration case only when Word and Zotero
are healthy and no valuable unsaved Word documents are open:

```powershell
.\.venv\Scripts\python.exe -m research_mcp.benchmark `
  --tiers read_only versioned zotero_live planned
```

The runner refuses a live tier when it detects any open document or any
`WINWORD.EXE` process other than the recorded project worker. Process-level
checking is required because COM may expose the empty worker while hiding a
separate Word instance with a valuable document. The
`--allow-open-word-documents` override is explicit risk acceptance and should
not be used for routine benchmarking.

Run one case or repeat for reliability estimates:

```powershell
.\.venv\Scripts\python.exe -m research_mcp.benchmark `
  --case search_zotero --repeat 5
```

Each run writes to a new `benchmark_runs/<timestamp>/` directory:

- `events.jsonl`: append-only event log after every model/tool step;
- `transcripts/`: complete per-case traces;
- `summary.json`: machine-readable metrics and verification results;
- `report.md`: compact success/token/latency table and all-vs-minimal deltas.

The runner reads `UNSLOTH_API_KEY` from the process environment or the project
`.env`. It defaults to a local OpenAI-compatible server at
`http://127.0.0.1:8888/v1` (`--base-url` and `--model` change it). Secrets are never written
to benchmark logs.

## Result meanings

- `passed`: the agent called the endpoint and the independent verifier observed
  the requested state, then stopped cleanly.
- `overrun`: the requested state verified, but the model continued making calls
  until the turn limit. The report separates tokens-to-verification from total
  tokens so this inefficiency remains visible.
- `passed_guarded`: a mutating task independently verified and the harness
  stopped immediately. This prevents redundant Word/Zotero writes; it does not
  claim that the model would have terminated cleanly on its own.
- `failed`: the model ran, but the requested state did not verify.
- `error`: provider, MCP, tool, or harness execution failed.
- `unsupported`: the current MCP does not expose a tool required by the case;
  this is a product capability gap, not a model failure.

Versioned document writes are confined to the run directory. Source documents
are never overwritten. The harness refuses any mutating MCP call without an
`output` path inside that directory. The DOI-import case is now executable, but
it is an external Zotero-library write and requires a disposable DOI and
cleanup plan. Batch citation insertion remains planned because its MCP tool
does not yet exist.

For live Word/Zotero cases, the runner records the project worker identity before
starting. It gracefully shuts down only a worker created during that benchmark
and verifies the exact Word PID plus creation-time identity exited. Pre-existing
workers and unrelated Word instances are never closed.
