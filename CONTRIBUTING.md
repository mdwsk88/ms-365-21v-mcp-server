# Contributing

Thank you for helping improve the project. Before public contributions are accepted, the maintainers must complete the license and corporate open-source approval gates in [Open-Source Release Plan](docs/OPEN_SOURCE_RELEASE.md).

## Development setup

```bash
npm ci
cp .env.example .env
npm run typecheck
npm test
```

Use placeholders or a dedicated non-production tenant for integration tests. Never commit `.env`, `.tokens`, private deployment overlays, real tenant/application IDs, access tokens, client secrets, internal domains, email addresses, or customer data.

## Change requirements

- Keep `direct` mode backward compatible.
- Dynamic discovery and execution must reuse `ToolRuntime`; do not create an authorization, confirmation, audit, or Graph-scope bypass.
- A new Graph tool must declare its delegated scopes, App Role category, read/write status, confirmation operation type, bounded inputs, and Chinese intent aliases.
- Claim 21V support only when the corresponding Microsoft Graph documentation explicitly supports China operated by 21Vianet for the selected API version and delegated permission model.
- Add focused tests for the changed behavior and run the complete suite.
- Update the generated tool catalog when tool metadata changes.

## Pull requests

Keep changes scoped and explain the security impact. Include test evidence, required Entra permissions, migration notes, and any behavior change to tool discovery or write confirmation.

Run before opening a pull request:

```bash
npm run check:public
npm test
npm audit --omit=dev
```
