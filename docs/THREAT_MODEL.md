# Threat Model

## Scope

This model covers the Streamable HTTP MCP server, OAuth bridge, Microsoft Entra token handling, OBO exchange, Graph tool runtime, dynamic tool routing, confirmation flow, audit logs, and deployment configuration. It does not replace the tenant's Conditional Access, data-classification, retention, or incident-response policies.

## Assets

- Microsoft user assertions, refresh/access tokens, MCP bridge tokens, and client credentials
- Mail, calendar, files, SharePoint, Teams, contacts, and directory data
- Entra App Role and Graph scope policy
- Confirmation tokens and pending write operations
- Audit records and operational metrics

## Trust boundaries

1. MCP client to public MCP endpoint
2. Browser and MCP OAuth bridge to Microsoft Entra
3. MCP server to Microsoft Graph China through OBO
4. Confirmation browser to the one-time approval endpoint
5. Process memory and protected local state to the host filesystem
6. Reverse proxy or managed platform to the Node.js process

## Principal threats and mitigations

| Threat | Mitigation |
|---|---|
| Missing, forged, wrong-audience, or wrong-scope MCP token | JWT signature, issuer, audience, expiry, and required-scope validation before tool calls |
| MCP token reused as a Graph token | Separate resource audience and OBO exchange; Graph receives only a Graph token |
| Cross-user session or token confusion | Per-request claims/context and user-bound bridge/confirmation state |
| Tool hidden from the catalog but invoked directly | App Roles and deployment policy are checked again by `ToolRuntime` |
| Dynamic executor bypasses policy | Search filters by current roles; schema and execute conceal unauthorized tools; execute re-enters `ToolRuntime` |
| Prompt injection triggers a destructive write | Deployment-level operation policy, MCP elicitation, or user-bound one-time browser approval |
| Agent self-approves a protected write | `confirm_execute` continues only a token already approved through the human channel |
| Confirmation replay or approval by another user | Short TTL, user binding, same-site browser cookie, and single-use token consumption |
| Secret or personal-data leakage in logs | Parameter redaction, bounded strings, protected audit directory, and file mode `0600` |
| Retry duplicates a write | Ambiguous POST/write failures are not automatically replayed |
| Graph throttling or outage cascades | Bounded timeout, `Retry-After`, exponential backoff for safe requests, and circuit breaker |
| Excessive tool context or model misrouting | Role/scope filtering plus optional discovery and hybrid exposure modes |
| Open redirect through OAuth client callbacks | Registered redirect validation and separate fixed Microsoft callback in bridge mode |
| Authorization-code phishing through a dynamically registered client (confused deputy: Microsoft shows no consent prompt for this server's already consented client) | Client redirect URIs limited to loopback addresses, native app schemes and https hosts listed in `MCP_OAUTH_REDIRECT_HOSTS`, checked at registration and again at every authorization; registration can be turned off |
| Request floods, runaway agents or registration spam | Per-address limits on the OAuth, attachment, confirmation and admin routes; per-user limits on the MCP endpoint; capped client registrations and MCP sessions |
| Mail attachment download link leaks | Random single-use ticket, lifetime of at most 280 seconds, per-user cap, `no-store`, `nosniff` and sandbox CSP headers; the Graph token never leaves the server |

## Residual risks

- A correctly authorized user can still ask an Agent to perform an unsafe action; confirmation reduces but cannot eliminate social engineering.
- `automatic` send mode deliberately removes human confirmation for email and Teams sends and should use a dedicated deployment and narrowly assigned roles.
- Microsoft Graph and the MCP client remain external trust dependencies.
- Conditional Access can permit or deny a login independently at the MCP resource and downstream Graph layers.
- Host compromise can expose process memory and locally stored OAuth state; production deployments need OS hardening, secret management, TLS termination, patching, and centralized monitoring.
- Dynamic client registration still lets anyone register a client whose redirect URI the policy allows; set `MCP_OAUTH_DYNAMIC_REGISTRATION=false` once every client is known in advance.
- Rate limits live in process memory and count per instance; multi-instance deployments and distributed floods need limits at the proxy or gateway as well.

## Security verification

The automated suite covers token boundaries, OAuth routing, role filtering, scope filtering, dynamic execution, confirmation replay/user binding, audit redaction, unsafe retry behavior, and HTTP transport behavior. Production releases also require a clean secret scan and a tenant-level sign-in/Graph smoke test.
