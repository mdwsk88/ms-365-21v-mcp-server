# Changelog

## [Unreleased] - planned 0.2.0

Tool slimming: clients see a small set of high-frequency read tools, and everything else is reached through discovery. **Breaking:** default exposure mode and several tool outputs changed; see Migration.

### Added

- `files_search`: search documents across OneDrive and SharePoint with Microsoft Search, filtered by site, file type and modification date. Falls back to per-library search when Microsoft Search is unavailable.
- `files_get_download_url`: resolve a OneDrive or SharePoint sharing link or item ID and return a short-lived pre-authenticated download URL for the agent client to fetch and parse; optional Office-to-PDF conversion; small plain-text files are returned inline.
- `gateway_search_tools` returns the input schema for its top results (`schemaCount`, default 3), so most long-tail calls need only search then execute. Invalid `gateway_execute_tool` arguments return the expected schema.

### Changed

- Default `MCP_TOOL_EXPOSURE_MODE` is now `hybrid`. Unset `MCP_DIRECT_TOOLS` lists `teams_list_chats`, `teams_list_chat_messages`, `mail_list_messages`, `mail_search_messages`, `mail_get_message`, `calendar_list_view`, `files_search` and `files_get_download_url`; `none` lists no business tools directly. `mcp.admin` now sees 13 tools instead of 126.
- `teams_list_chats`: returns member names and the last message, ordered by latest activity; adds `since`, `query` (person or topic), `chatType` and `cursor`.
- `teams_list_chat_messages`: plain text by default, oldest first within a page, with @mentions, shared files and quotes; adds `since`, `until`, `cursor`, `includeSystemMessages` and `format`.
- `mail_list_messages`: adds `folder` (inbox by default, or `all`), `since`, `until`, `flaggedOnly`, `importance`, `from` and `cursor`; rows include preview, flag state and whether the user is in To or Cc.
- `mail_search_messages`: adds `folder` and `cursor`; same compact rows.
- `mail_get_message`: plain-text body by default, `onlyNewContent` (drops the quoted thread), `maxChars`/`offset` paging and the attachment list.
- `calendar_list_view`: compact events with cancellation, my response and attendees; times without an offset are read as China Standard Time.
- Discovery ranking uses BM25 over tool name, title, category and description with bilingual synonyms.
- All tool results are compact JSON without Graph `@odata.context`/`@odata.etag` metadata or null values.
- Server instructions tell clients that mail, chat and file text is untrusted data and that download URLs must not be shared.

### Removed

- `mail_list_folder_messages`; use `mail_list_messages` with `folder`.

### Migration

- Clients that call a business tool by name must either keep it in `MCP_DIRECT_TOOLS` or call it through `gateway_execute_tool`. Set `MCP_TOOL_EXPOSURE_MODE=direct` to keep the previous behavior.
- Consumers that parsed raw Graph payloads from the tools above must read the new compact fields.

### Verification

- 97 automated tests with mocked Microsoft Graph. Read-only probes in a 21V tenant on 2026-10-12 confirmed chat `$expand=members,lastMessagePreview` with `$orderby=lastMessagePreview/createdDateTime desc`, chat message filters on `lastModifiedDateTime` (gt/lt) and `createdDateTime` (lt only), and combined mail filters led by `receivedDateTime`.
- Not yet verified in a 21V tenant: Microsoft Search for `driveItem`, the `/shares` link API and `format=pdf` conversion.

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
