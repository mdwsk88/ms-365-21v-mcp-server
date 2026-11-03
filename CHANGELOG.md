# Changelog

## [Unreleased]

Maintenance release: mail attachments as download links, batch and cleaned mail reading, shared mailboxes and calendars, a read-only mode with a Graph permission allow-list, and hardening of the OAuth bridge and HTTP endpoint. **Breaking:** the smart tools are retired, OAuth client redirect URIs are restricted, and several outputs changed; see Migration.

### Security

- The OAuth bridge accepted any redirect URI at dynamic registration. Microsoft shows no consent prompt for the server's own, already consented client, so anyone could register a client whose redirect URI they control and phish a signed-in user's authorization code. Redirect URIs must now be a loopback address, a native app scheme, or https on a host listed in `MCP_OAUTH_REDIRECT_HOSTS`. The policy applies at registration and at every authorization, so clients registered earlier are covered too.

### Added

- `mail_get_attachment` returns small text attachments inline, attached cloud files as their source URL, and other files as a single-use download URL served by this server at `/attachments/{ticket}`. A link expires after `MCP_ATTACHMENT_URL_TTL_SECONDS` (240, at most 280) and works once; the Graph token never leaves the server. stdio deployments keep base64 (up to 10 MB). The tool joins the default direct set.
- `mail_get_message` reads up to 10 messages in one call with `messageIds`; a failing message comes back with its Graph status and error code.
- `mailbox` on `mail_list_messages`, `mail_search_messages`, `mail_get_message` and `mail_get_attachment` reads a shared or delegated mailbox (needs `Mail.Read.Shared`). `calendarOwner` on `calendar_list_view` reads a calendar shared with the user (needs `Calendars.Read.Shared`), and `cursor` continues a long calendar view.
- `calendar_list_view` flags overlapping meetings with `overlapsWith` and `conflictCount`, ignoring cancelled, declined, all-day and free events.
- `MCP_READ_ONLY=true` removes every tool that sends, creates, changes or deletes data, together with `confirm_execute`.
- `MCP_ALLOWED_GRAPH_SCOPES` lists the consented delegated permissions; tools that need anything else are hidden (fail closed). Broader scopes cover narrower ones. `MCP_DISABLED_GRAPH_SCOPES` still wins.
- `npm run permissions` (`node dist/index.js --list-permissions [--json]`) prints the scopes the enabled tools need, optional per-input scopes, hidden tools with their missing scopes, and allowed scopes no tool uses, from configuration alone.
- `MCP_OAUTH_DYNAMIC_REGISTRATION=false` turns client registration off. Registration keeps at most 1000 clients and 10 redirect URIs per client.
- In-memory rate limits, answered with 429 and `Retry-After`: per client address on the OAuth, attachment, confirmation and admin routes (`MCP_RATE_LIMIT_IP_PER_MINUTE`, 120; registration a tenth of it), and per signed-in user on `/mcp` (`MCP_RATE_LIMIT_USER_PER_MINUTE`, 600). `MCP_TRUST_PROXY` decides which `X-Forwarded-For` hops count; a loopback bind trusts its same-host proxy by default.
- Stateful MCP sessions close after `MCP_HTTP_SESSION_IDLE_MINUTES` (480) idle minutes, and at most `MCP_HTTP_MAX_SESSIONS` (200) are kept. Before, every session stayed in memory, about 2.6 MB each, until the client sent DELETE.
- `/healthz` reports `readOnly`, `allowedGraphScopes` and `dynamicClientRegistration`.
- Release automation publishes `server.json` to the official MCP Registry after each successful tagged release, using GitHub Actions OIDC. It skips versions that are already listed and can be run manually for an existing tag.
- A documentation test keeps `server.json` and the Dockerfile ownership label in step with the package version.

### Changed

- Mail text bodies are cleaned by default: Safe Links unwrapped (any cloud, including 21Vianet), tracking parameters and repeated link targets dropped, inline images shown as `[image]`, and links over 500 characters outside Microsoft 365 shortened to `[link: host]`. `clean: false` returns the text as Outlook renders it. Mail previews are cleaned the same way, and every text the server returns loses zero-width and bidi-control padding.
- `search_teams` returns one compact row per hit (sender, time, subject, summary, and the chat or team and channel to read next) with `total` and `nextFrom`. It stays out of the default direct tools until chat message search is verified in 21Vianet.
- `teams_list_channel_messages` and `teams_list_channel_message_replies` return the chat message shape plus `subject` and `webLink`, skip deleted posts, sort each page oldest first and accept `cursor`.
- Opening a file by link (`files_get_download_url` with `url`) requests `Files.ReadWrite`, the least-privileged permission of the shares API, and fails with `scope_not_enabled` where it is not allowed. PDF conversion of files over 2 MB returns the original file with a note instead of failing.
- `users_list_direct_reports` needs only `User.Read` and `User.ReadBasic.All`, and `users_list_memberships` only `User.Read` for the signed-in user, so both stay available without `User.Read.All`.
- A refresh token presented with another client's `client_id` is refused.

### Fixed

- stdio transport: dotenv 17 printed a log line to stdout on every start, so every stdio session since v0.1.0, including `docker run -i` from the MCP Registry entry, began with a line that is not JSON-RPC.
- Text files read inline from attachments, OneDrive or SharePoint kept only single spaces and lost tabs, indentation and blank lines, which broke YAML, Python and TSV content. Only line endings and a byte-order mark change now.

### Removed

- `smart_mail_digest`, `smart_calendar_conflicts` and `smart_teams_unread`, the `smart` category and the `mcp.smart` App Role in the role template. Use `mail_list_messages`, `calendar_list_view` and `teams_list_chats` with `since`.
- The `raw` parameter of `mail_get_attachment`; use `format: "base64"` for inline bytes.

### Migration

- Before upgrading an OAuth bridge deployment, list the hosts of web-based clients with https redirect URIs in `MCP_OAUTH_REDIRECT_HOSTS`; the `dynamic_client_registered` log events show each client's redirect origin. Local clients on loopback addresses need nothing.
- Clients that called the smart tools, parsed raw payloads from `search_teams` or the channel message tools, or relied on `mail_get_attachment` returning base64 must switch to the replacements above.
- An existing `.env` that sets `MCP_DIRECT_TOOLS` explicitly keeps its list; add `mail_get_attachment`, or clear the variable to use the new default set.
- No Microsoft Entra change is required. `Mail.Read.Shared`, `Calendars.Read.Shared`, `ChannelMessage.Read.All` and `User.Read.All` enable optional features once they are consented and allowed. The `mcp.smart` role can stay assigned without effect, or be disabled and removed.

## [0.2.0] - 2026-10-14

Tool slimming: clients see a small set of high-frequency read tools, and everything else is reached through discovery. Also adds Host header validation for loopback deployments behind a reverse proxy. **Breaking:** default exposure mode and several tool outputs changed; see Migration.

### Added

- `files_search`: search documents across OneDrive and SharePoint with Microsoft Search, filtered by site, file type and modification date. Falls back to per-library search when Microsoft Search is unavailable.
- `files_get_download_url`: resolve a OneDrive or SharePoint sharing link or item ID and return a short-lived pre-authenticated download URL for the agent client to fetch and parse; optional Office-to-PDF conversion; small plain-text files are returned inline.
- `gateway_search_tools` returns the input schema for its top results (`schemaCount`, default 3), so most long-tail calls need only search then execute. Invalid `gateway_execute_tool` arguments return the expected schema.
- `MCP_ALLOWED_HOSTS` and a Host header allow-list against DNS rebinding. A loopback bind now accepts `localhost`, `127.0.0.1`, `[::1]` and the configured public host names (resource URL, public base URL, OAuth issuer), so a same-host reverse proxy that forwards the public `Host` header works with `MCP_HTTP_HOST=127.0.0.1`. Previously the MCP SDK accepted only localhost names on a loopback bind, which forced such deployments onto `0.0.0.0`. `MCP_ALLOWED_HOSTS` adds names and also enables the check for other binds; an all-interfaces bind without it is unchanged.
- `/healthz` reports the server `version`.

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
- `files_search` and `files_get_download_url` need `mcp.drive`; SharePoint files additionally need `mcp.sharepoint`, so the new tools keep the existing OneDrive/SharePoint App Role boundary. A caller without `mcp.sharepoint` gets OneDrive results only and a `permission_denied` error for SharePoint sites and files.

### Removed

- `mail_list_folder_messages`; use `mail_list_messages` with `folder`.

### Migration

- Clients that call a business tool by name must either keep it in `MCP_DIRECT_TOOLS` or call it through `gateway_execute_tool`. Set `MCP_TOOL_EXPOSURE_MODE=direct` to keep the previous behavior.
- Consumers that parsed raw Graph payloads from the tools above must read the new compact fields.

### Verification

- 102 automated tests with mocked Microsoft Graph, including a spawned loopback server that accepts the public Host header and answers 403 to others. Read-only probes in a 21V tenant on 2026-10-12 confirmed chat `$expand=members,lastMessagePreview` with `$orderby=lastMessagePreview/createdDateTime desc`, chat message filters on `lastModifiedDateTime` (gt/lt) and `createdDateTime` (lt only), and combined mail filters led by `receivedDateTime`.
- Not yet verified in a 21V tenant: Microsoft Search for `driveItem`, the `/shares` link API and `format=pdf` conversion.
- No Microsoft Entra changes are required: no new App Roles, delegated Graph permissions or redirect URIs.

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

[0.2.0]: https://github.com/mdwsk88/ms-365-21v-mcp-server/releases/tag/v0.2.0
[0.1.0]: https://github.com/mdwsk88/ms-365-21v-mcp-server/releases/tag/v0.1.0
