# Maintenance priorities

[Home](../README.md) · [Changelog](../CHANGELOG.md)

This is a proposed maintenance backlog, not a list of shipped features or a promise of delivery dates. Keep small, reviewable changes and validate them before publishing compatibility claims.

## 1. Dependency security and repeatable validation

PRs #8/#9 (`fast-uri` 3.1.7 and `qs` 6.16.0) were reviewed and merged through #10 on 2026-09-05. Continue reviewing other Dependabot pull requests against upstream advisories, dependency paths and full CI before merging.

Keep TypeScript major upgrades separate from onboarding and authentication changes. Align the supported Node runtime, Node type definitions, Docker image and CI matrix deliberately. The new onboarding matrix covers the standalone setup/doctor helpers, not the entire server on Windows.

Acceptance: a clean lockfile install, full test run, production dependency audit and documented upgrade/rollback notes. Preserve the existing public-release credential checks.

## 2. Live onboarding and deployment checks

Run the documented path on a clean Windows desktop and Linux host using a dedicated test tenant. Record client version, date, auth mode and which calls succeeded. Test Docker token-directory ownership, HTTPS proxy prefixes and service restart persistence before advertising one-command deployment.

Acceptance: login, auth_status, graph_get_me and one read-only mail query succeed; unauthenticated tool execution still fails; no secrets appear in diagnostics. Test the same docs with someone unfamiliar with the project.

## 3. Evidence-backed discovery and distribution

Create a short, redacted real demonstration: connect a client, log in, list a few messages, and show the confirmation boundary for a write. Never use real employee messages or tenant credentials in demo assets. Do not portray a mock recording as a live integration.

Versioned releases and a tested container publishing pipeline are included in v0.1.0. See [container distribution](CONTAINER.md). Continue recording real tenant onboarding separately from container smoke checks. npm/npx distribution is not published.

Acceptance: a newcomer understands the China-cloud use case from the README, reaches their first successful read, and can reproduce the demo. Useful measures are onboarding failure categories and time-to-first-success from consenting testers, not just star count.

## 4. Community feedback

Use the setup/bug and feature templates to capture reproducible reports without collecting tenant secrets. Keep the Chinese quickstart and English overview aligned. Turn verified fixes into changelog entries and document versioned compatibility results.

A star invitation is appropriate after showing real value; fabricated benchmarks, endorsements or compatibility badges are not.
