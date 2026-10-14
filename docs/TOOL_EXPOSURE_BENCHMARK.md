# Tool Exposure Benchmark

Run:

```bash
npm run benchmark:tools
```

The benchmark creates in-memory MCP clients, initializes a server, and serializes `tools/list`. It uses the production Graph-scope boundary from `.env.example`. It does not call a model, OAuth provider, network, or Microsoft Graph.

Reference result from 2026-10-12 on arm64 macOS with Node.js 26.7.0:

| Role profile | Mode | Exposed tools | Serialized `tools/list` |
|---|---|---:|---:|
| `mcp.admin` | direct | 126 | 100,438 bytes |
| `mcp.admin` | discovery | 5 | 3,890 bytes |
| `mcp.admin` | hybrid | 13 | 14,879 bytes |
| `mcp.mail` | direct | 24 | 20,220 bytes |
| `mcp.mail` | discovery | 5 | 3,890 bytes |
| `mcp.mail` | hybrid | 8 | 7,927 bytes |

The hybrid profile used the built-in default direct set (`defaultDirectTools`: Teams chats, mail, calendar view, file search and download URL). Counts change with Graph-scope and App Role policy. The scenario tools carry richer filters (time windows, flags, cursors), so each costs about 1.1-1.6 KB of schema; that trade keeps common requests to one or two calls.

The previous reference (2026-08-09, before the scenario tools) was 125 / 94,065 bytes for direct and 12 / 7,757 bytes for hybrid with a smaller, generic direct set.

## Interpretation

- Direct mode minimizes MCP/Agent tool-call rounds for a clear request because the target schema is already visible.
- Discovery mode reduces the initial tool schema by routing long-tail work through search, schema lookup, and execute calls.
- Hybrid mode keeps the direct fast path for the high-frequency read scenarios while cutting the admin catalog payload by about 85% in this configuration.
- Local initialization time is intentionally not used as a product latency claim; model provider, client orchestration, OAuth, Graph, network, query shape, and response size dominate real end-to-end measurements.

An end-to-end benchmark should use the same client, model, account, prompt set, warm/cold state, Graph query parameters, and result limits. Record model turns, MCP calls, Graph requests, time to first tool, total latency, response bytes, and task success.
