# Open-Source Release Process

## Current status

The project owner has approved publishing the reusable core as open source. The project uses Apache License 2.0. The existing private Git history previously contained deployment-specific tenant/application IDs and service URLs even though the current working tree is sanitized.

The public repository was initialized from a sanitized, history-free snapshot. Keep the original private GitHub/GitLab repositories private and never merge their history into the public repository. Subsequent versions are released from the public repository's reviewed `main` branch.

## Versioned releases

1. Create a release branch from the public repository's latest `main`. Update `package.json` and `package-lock.json` together when changing the version, add `docs/releases/vX.Y.Z.md`, and update the changelog and default tag in `docker-compose.release.yml`.
2. Open a pull request. Require CI, the onboarding matrix and the Release container smoke job to pass, then merge it.
3. Tag the merged commit with the matching `vX.Y.Z` version and push that tag. The Release workflow checks the tag against `package.json`, runs the test/audit/public-release gates, and verifies the container's protected HTTP endpoint and state persistence.
4. After verification, the workflow publishes Linux amd64/arm64 images to GHCR with version and `latest` tags, SBOM/provenance metadata, then creates the GitHub Release with deployment assets and checksums. It publishes only from the public repository. It does not deploy to AWS/AgentRun or publish to npm.
5. On first publication, change the new GHCR package's visibility to public in GitHub Package settings. Test an anonymous image pull and repeat the container smoke check against the published image. Record the release URL and image digest.

Keep version tags fixed once published. Use a new patch version for corrections. For a failed workflow, fix the cause before rerunning; `workflow_dispatch` on an existing version tag can resume publication. A dispatch on a branch only runs verification. Preserve existing deployment configuration and state during upgrades; see [container upgrade instructions](CONTAINER.md).

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

The project should be positioned as a 21V-first, policy-aware Microsoft 365 MCP gateway with a direct Chinese fast path and optional long-tail discovery. It should not claim affiliation with Microsoft or 21Vianet, or present itself as an official edition of another project.

Tool count alone is not the goal. Public benchmarks should compare direct, discovery, and hybrid modes by tool-schema bytes, model/tool round trips, Graph requests, latency, result size, and task success for Chinese enterprise prompts.
