#!/usr/bin/env bash
set -euo pipefail

project_root="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
output_path="${1:-$HOME/Desktop/ms-365-21v-mcp-server-agentrun.zip}"
staging_dir="$(mktemp -d "${TMPDIR:-/tmp}/ms-365-21v-mcp-server-agentrun.XXXXXX")"

cleanup() {
  rm -rf "$staging_dir"
}
trap cleanup EXIT

cd "$project_root"
npm run build

cp package.json package-lock.json "$staging_dir/"
cp -R dist "$staging_dir/"
(
  cd "$staging_dir"
  npm ci --omit=dev --ignore-scripts
)

mkdir -p "$(dirname "$output_path")"
rm -f "$output_path"
(
  cd "$staging_dir"
  zip -qr "$output_path" .
)

printf 'AgentRun package: %s\n' "$output_path"
