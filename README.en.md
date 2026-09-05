# MS 365-21V MCP Server

**Connect MCP-capable AI clients to Microsoft 365 operated by 21Vianet, as the signed-in user.**

[中文](README.md) · [Quickstart (Chinese)](docs/QUICKSTART.md) · [Deployment reference (Chinese)](DEPLOYMENT.md) · [Tool catalog](docs/TOOL_CATALOG.md)

A Streamable HTTP MCP gateway with an OAuth bridge, delegated Microsoft Graph access, per-user App Roles, write confirmations and audit logging. It targets the China cloud, not global Microsoft 365 or personal Outlook accounts. This is not an official Microsoft or 21Vianet product.

## Start with one successful read

You need a 21V tenant, administrator approval and Node.js 22+. Register a single-tenant MCP API application with a Web callback, expose `access_as_user`, grant Graph delegated `User.Read`, and assign `mcp.users` to your test user. Follow the [setup checklist](docs/QUICKSTART.md); the helper does not create or change Entra applications.

```bash
git clone https://github.com/mdwsk88/ms-365-21v-mcp-server.git
cd ms-365-21v-mcp-server
npm run setup
```

The helper asks for the tenant ID, API application ID and public base URL. It refuses to overwrite `.env`. Add the client secret **value** to that file, not to shell arguments, then run:

```bash
npm run doctor
npm ci
npm run build
npm run start:http
```

Use `http://localhost:3000/mcp` only for a local desktop client. Remote/cloud clients need a reachable HTTPS endpoint. Complete login, call `auth_status`, then `graph_get_me`. A healthy HTTP process alone is not proof of working Graph authorization.

The generated profile keeps authentication, App Roles, audit logging and confirmation enabled. It loads the profile module first. Expand to mail, calendar and other modules only after granting the required permissions and roles.

Already have a deployment? Use `npm run doctor`; do not replace an existing `.env`. Docker and two-application setups remain documented in [DEPLOYMENT.md](DEPLOYMENT.md).

## Examples after enabling the relevant modules

| Ask your client | Required capability |
|---|---|
| List my five latest emails, showing only subject and sender | Mail read |
| Check tomorrow's calendar for overlapping meetings | Calendar + Smart |
| Find the project plan in my OneDrive | OneDrive read |
| Prepare an email reply and ask before sending | Mail write + confirmation |

Modules include mail, calendar, OneDrive, SharePoint, Teams, contacts, users, Microsoft Search and aggregation tools. Availability depends on tenant consent, App Roles, configured filters and 21V API support. See the [scope-to-tool mapping](docs/TOOL_CATALOG.md).

## Diagnostics and trust

`npm run doctor` checks common HTTP/OAuth configuration errors offline. For clean JSON use `npm run --silent doctor -- --json`; use `--config PATH` to check another file. Process environment values take precedence. Output does not contain configuration values.

This is not a live login test or a production security certification. Verify consent, Conditional Access, role assignments, proxy routing and Graph calls separately. Keep credentials and token caches out of source control.

China Graph endpoints do not guarantee that data stays in China: results are also supplied to your AI client/model provider. Review that data flow before production use.

The repository previously recorded WorkBuddy, Qoder Work, Codex and Dify interoperability. This maintenance change does not re-certify their latest versions. Clients need Streamable HTTP, OAuth discovery, browser callbacks and Bearer token support.

## Contribute

Read [CONTRIBUTING.md](CONTRIBUTING.md), use the issue templates for redacted setup/compatibility reports, and report vulnerabilities through [SECURITY.md](SECURITY.md), not public issues. A star helps others discover the project when it solves a real 21V integration need.

[Maintenance priorities](docs/MAINTENANCE.md) · [Changelog](CHANGELOG.md) · [Apache-2.0 license](LICENSE) · [Trademarks](TRADEMARKS.md)
