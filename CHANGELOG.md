# Changelog

## Unreleased

### Added

- Dependency-free `npm run setup` for a minimal single-app HTTP/OAuth configuration, with no secret command-line input and no overwrite of an existing `.env`.
- Offline `npm run doctor` with structured JSON, explicit error exit codes, common China-cloud/OAuth configuration checks and value-free diagnostic messages.
- Onboarding unit and CLI regression tests, plus a dedicated Linux/Windows, Node 22/24 CI matrix for those helpers.
- Chinese guided quickstart, English overview, maintenance priorities and issue templates.

### Changed

- Put user scenarios and the first successful profile read at the front of the README; preserve the previous detailed README as `DEPLOYMENT.md`.
- Use the existing `--http` flag in npm HTTP commands and Node file permissions instead of a shell `chmod` command in the build script.
- Leave `MCP_RESOURCE_URL` empty in `.env.example` so an example hostname no longer overrides the configured public base URL. Existing `.env` files are not changed automatically.

### Scope

- No new runtime dependencies, package release, cloud deployment or authentication-policy relaxation.
- Live 21V consent, Conditional Access and client interoperability must be validated in the target tenant. Offline preflight is not a substitute.
