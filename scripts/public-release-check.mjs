#!/usr/bin/env node
import fs from 'node:fs';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';

const root = fileURLToPath(new URL('..', import.meta.url));
const releaseMode = process.argv.includes('--release');
const failures = [];

function projectFiles() {
  try {
    const output = execFileSync(
      'git',
      ['ls-files', '--cached', '--others', '--exclude-standard', '-z'],
      { cwd: root, stdio: ['ignore', 'pipe', 'ignore'] }
    ).toString('utf8');
    return output.split('\0').filter(Boolean);
  } catch {
    const excludedDirectories = new Set(['.git', '.tokens', 'dist', 'node_modules', 'private']);
    const excludedFiles = new Set(['.env', '.public-release-deny-patterns']);
    const files = [];
    const visit = directory => {
      for (const entry of fs.readdirSync(directory, { withFileTypes: true })) {
        if (entry.isDirectory() && excludedDirectories.has(entry.name)) continue;
        const absolute = path.join(directory, entry.name);
        const relative = path.relative(root, absolute);
        if (entry.isDirectory()) visit(absolute);
        else if (entry.isFile() && !excludedFiles.has(entry.name)) files.push(relative);
      }
    };
    visit(root);
    return files;
  }
}

function customDenyPatterns() {
  const file = path.join(root, '.public-release-deny-patterns');
  if (!fs.existsSync(file)) return [];
  return fs
    .readFileSync(file, 'utf8')
    .split(/\r?\n/)
    .map(value => value.trim())
    .filter(value => value && !value.startsWith('#'));
}

function isText(contents) {
  return !contents.subarray(0, 8_192).includes(0);
}

function record(file, message) {
  failures.push(`${file}: ${message}`);
}

const files = projectFiles();
const forbiddenTrackedPaths = [
  /^\.env$/,
  /^\.tokens(?:\/|$)/,
  /^\.codex\/.*\.env$/,
  /^deploy\/overlays\/private(?:\/|$)/,
  /^\.public-release-deny-patterns$/
];

for (const file of files) {
  if (forbiddenTrackedPaths.some(pattern => pattern.test(file))) {
    record(file, 'private deployment material must not be tracked');
    continue;
  }
  const absolute = path.join(root, file);
  let contents;
  try {
    contents = fs.readFileSync(absolute);
  } catch {
    continue;
  }
  if (contents.length > 2_000_000 || !isText(contents)) continue;
  const text = contents.toString('utf8');
  const checks = [
    [/-----BEGIN (?:RSA |EC |OPENSSH )?PRIVATE KEY-----/, 'private key material detected'],
    [/\b(?:AKIA|ASIA)[A-Z0-9]{16}\b/, 'AWS access key ID detected'],
    [/\bgh[pousr]_[A-Za-z0-9_]{30,}\b/, 'GitHub token detected'],
    [/\bglpat-[A-Za-z0-9_-]{20,}\b/, 'GitLab token detected'],
    [
      /^(?:MS_CLIENT_SECRET|MS_OAUTH_CLIENT_SECRET|MCP_ADMIN_TOKEN)=(?!\s*$|<|your-|replace-).+/im,
      'non-placeholder secret assignment detected'
    ],
    [
      /^(?:MS_TENANT_ID|MS_CLIENT_ID|MS_PUBLIC_CLIENT_ID|MS_OAUTH_CLIENT_ID)=(?!0{8}-0{4}-0{4}-0{4}-0{12})[0-9a-f]{8}-[0-9a-f-]{27}$/im,
      'real tenant or application UUID detected in a tracked config file'
    ]
  ];
  for (const [pattern, message] of checks) {
    if (pattern.test(text)) record(file, message);
  }
  for (const denied of customDenyPatterns()) {
    if (text.toLowerCase().includes(denied.toLowerCase())) {
      record(file, `matches private deny pattern ${JSON.stringify(denied)}`);
    }
  }
}

if (releaseMode) {
  for (const required of ['LICENSE', 'SECURITY.md', 'CONTRIBUTING.md']) {
    if (!fs.existsSync(path.join(root, required))) {
      record(required, 'required before creating a public repository');
    }
  }
}

if (failures.length > 0) {
  console.error(`Public release check failed with ${failures.length} issue(s):`);
  for (const failure of failures) console.error(`- ${failure}`);
  process.exitCode = 1;
} else {
  console.log(`Public release ${releaseMode ? 'gate' : 'preparation'} check passed (${files.length} files scanned).`);
}
