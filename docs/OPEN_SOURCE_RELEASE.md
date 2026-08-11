# Open-Source Release Process

## Current status

The project owner has approved publishing the reusable core as open source. The project uses Apache License 2.0. The existing private Git history previously contained deployment-specific tenant/application IDs and service URLs even though the current working tree is sanitized.

Do not change the visibility of the existing GitHub or GitLab repository. Public releases must come from a sanitized, history-free snapshot in a separate public repository.

## Repository boundary

### Public core

- OAuth 2.1 protected resource and 21V Entra bridge
- OBO token exchange and Graph China client
- direct/discovery/hybrid tool routing
- App Role, Graph scope, confirmation, audit, metrics, and resilience policies
- smart aggregation tools, generic deployment examples, tests, and documentation

### Private deployment overlay

- real tenant, application, subscription, account, instance, and role-assignment identifiers
- client secrets, certificates, admin tokens, AWS/AgentRun credentials, and private DNS names
- organization-specific Conditional Access decisions, user/group assignments, runbooks, and support contacts

Keep private values in a separate private repository and secret manager. The public repository contains only [the generic overlay contract](../deploy/overlays/README.md).

## Mandatory gates

1. Confirm owner and organization approval for the intended release.
2. Keep the Apache-2.0 `LICENSE` and any later required `NOTICE` file in the release.
3. Rotate every credential that has ever appeared in chat, screenshots, terminals, packages, or deployment experiments.
4. Run a full-history secret scanner on the private repository and investigate every finding.
5. Run `npm run check:public`, `npm test`, and `npm audit --omit=dev`.
6. Run `npm run check:release`.
7. Export a history-free tree with `npm run export:public -- /absolute/empty/path`.
8. Review the exported tree and initialize a new Git repository there.
9. Enable branch protection, dependency updates, private vulnerability reporting, and secret scanning on the public repository.

## Positioning

The project should be positioned as a 21V-first, policy-aware Microsoft 365 MCP gateway with a direct Chinese fast path and optional long-tail discovery. It should not claim affiliation with Microsoft or 21Vianet and should not market itself as a fork or official "Pro" edition of another project.

Tool count alone is not the goal. Public benchmarks should compare direct, discovery, and hybrid modes by tool-schema bytes, model/tool round trips, Graph requests, latency, result size, and task success for Chinese enterprise prompts.
