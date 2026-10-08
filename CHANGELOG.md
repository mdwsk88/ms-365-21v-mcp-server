# Changelog

## [0.1.0] - 2026-09-07

First versioned release, including the onboarding and security updates merged on 2026-09-05.

### Added

- Versioned GitHub Release and GHCR container distribution for Linux amd64/arm64, with deployment assets, image digest and checksums.
- Container smoke verification for health, OAuth challenge/discovery, non-root state access and restart persistence.
- Prebuilt-image Compose configuration using a named state volume, with installation and rollback instructions.
- Dependency-free `npm run setup` for a minimal single-app HTTP/OAuth configuration, with no secret command-line input and no overwrite of an existing `.env`.
- Offline `npm run doctor` with structured JSON, explicit error exit codes, common China-cloud/OAuth configuration checks and value-free diagnostic messages.
- Onboarding unit and CLI regression tests, plus a dedicated Linux/Windows, Node 22/24 CI matrix for those helpers.
- Chinese guided quickstart, English overview, maintenance priorities and issue templates.

### Changed

- Put user scenarios and the first successful profile read at the front of the README; preserve the previous detailed README as `DEPLOYMENT.md`.
- Use the existing `--http` flag in npm HTTP commands and Node file permissions instead of a shell `chmod` command in the build script.
- Leave `MCP_RESOURCE_URL` empty in `.env.example` so an example hostname no longer overrides the configured public base URL. Existing `.env` files are not changed automatically.

### Scope

- Includes reviewed `fast-uri` 3.1.7 and `qs` 6.16.0 dependency security updates from PRs #8/#9, merged through #10.
- No new runtime dependencies, npm publication, production cloud deployment or authentication-policy relaxation.
- Live 21V consent, Conditional Access and client interoperability must be validated in the target tenant. Offline preflight is not a substitute.

[0.1.0]: https://github.com/mdwsk88/ms-365-21v-mcp-server/releases/tag/v0.1.0
