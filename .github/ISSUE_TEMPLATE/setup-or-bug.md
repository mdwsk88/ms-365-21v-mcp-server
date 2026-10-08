---
name: Setup problem or bug
about: Report a reproducible setup, login or tool issue with redacted diagnostics
---

## What happened?

Expected result and actual result:

## Environment

- Repository commit or release:
- OS and Node/Docker version:
- MCP client and version:
- Local desktop or remote/cloud client:
- Single-app or two-app login; public or confidential_web:
- Transport and deployment type (do not include private hostnames):

## Steps to reproduce

1.
2.
3.

## Diagnostic results

Run `npm run --silent doctor -- --json` and review the result before sharing it.
State whether healthz, auth_status and graph_get_me work. Include only relevant, redacted errors.

## Privacy and security

- [ ] I have removed secrets, tokens, authorization headers, tenant/application IDs, personal data and private URLs.
- [ ] I am not attaching `.env`, token-cache files, full login URLs or employee messages.
- [ ] This is not a security vulnerability; I will use SECURITY.md for private vulnerability reporting.
