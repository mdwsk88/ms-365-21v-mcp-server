# Tool Exposure Benchmark

Run:

```bash
npm run benchmark:tools
```

The benchmark creates in-memory MCP clients, initializes a server, and serializes `tools/list`. It uses the production Graph-scope boundary from `.env.example`. It does not call a model, OAuth provider, network, or Microsoft Graph.

Reference result from 2026-08-09 on arm64 macOS with Node.js 26.7.0:

| Role profile | Mode | Exposed tools | Serialized `tools/list` |
|---|---|---:|---:|
| `mcp.admin` | direct | 125 | 94,065 bytes |
| `mcp.admin` | discovery | 5 | 3,550 bytes |
| `mcp.admin` | hybrid | 12 | 7,757 bytes |
| `mcp.mail` | direct | 25 | 18,773 bytes |
| `mcp.mail` | discovery | 5 | 3,550 bytes |
| `mcp.mail` | hybrid | 7 | 4,845 bytes |

The hybrid profile used `smart` as a direct category plus seven common direct tool names. Counts change with Graph-scope and App Role policy.

## Interpretation

- Direct mode minimizes MCP/Agent tool-call rounds for a clear request because the target schema is already visible.
- Discovery mode reduces the initial tool schema by routing long-tail work through search, schema lookup, and execute calls.
- Hybrid mode keeps the direct fast path while cutting the admin catalog payload by more than 90% in this configuration.
- Local initialization time is intentionally not used as a product latency claim; model provider, client orchestration, OAuth, Graph, network, query shape, and response size dominate real end-to-end measurements.

An end-to-end benchmark should use the same client, model, account, prompt set, warm/cold state, Graph query parameters, and result limits. Record model turns, MCP calls, Graph requests, time to first tool, total latency, response bytes, and task success.
