# Deployment Overlays

This directory defines the boundary between the public core and private production configuration.

- `example.env` documents the public contract with placeholders.
- `private/` is ignored and may be used for local experiments only.
- Real production overlays should live in a separate private repository and secret manager, not in this source tree.

An overlay may select login type, tool exposure mode, fast-lane tools, allowed Graph scope boundary, confirmation policy, and public routing. It must never contain a secret in a public commit.
