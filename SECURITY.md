# Security Policy

## Supported versions

Security fixes are applied to the latest release and the default branch. Pre-release branches are not supported production releases.

## Reporting a vulnerability

Do not open a public issue for authentication bypasses, token disclosure, cross-user data access, unsafe write execution, or other vulnerabilities.

Use GitHub Private Vulnerability Reporting for the public repository. Until that channel is enabled, report the issue through the maintainers' private organizational security channel. Include the affected version, deployment mode, reproduction steps, expected impact, and sanitized logs. Never include live tokens, client secrets, tenant identifiers, message bodies, or files.

The maintainers should acknowledge a report within five business days, validate severity, prepare a coordinated fix, and publish a security advisory when users must upgrade or rotate credentials.

## Security boundaries

- The MCP authenticates a user to the protected MCP resource, then uses OBO to obtain a separate Microsoft Graph token.
- Microsoft Entra Conditional Access, tenant consent, Graph delegated permissions, and user access to the target resource remain external enforcement layers.
- Entra App Roles filter both tool discovery and invocation. Dynamic tool execution re-enters the same invocation guard.
- Sensitive writes can require client-native or browser confirmation. Automatic sending is a deployment policy and does not bypass authorization or audit.
- Production secrets belong in a secret manager or protected environment file, never source control.

See [Threat Model](docs/THREAT_MODEL.md) for assumptions, mitigations, and residual risks.
