#!/usr/bin/env node
import fs from 'node:fs';
import path from 'node:path';
import { execFileSync, spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';

const root = fileURLToPath(new URL('..', import.meta.url));
const targetArgument = process.argv[2];
if (!targetArgument) {
  console.error('Usage: npm run export:public -- /absolute/path/to/empty-target');
  process.exit(2);
}

const target = path.resolve(targetArgument);
if (target === root || target.startsWith(`${root}${path.sep}`)) {
  console.error('The public snapshot target must be outside the source repository.');
  process.exit(2);
}
if (fs.existsSync(target) && fs.readdirSync(target).length > 0) {
  console.error(`Target is not empty: ${target}`);
  process.exit(2);
}

const check = spawnSync(process.execPath, [path.join(root, 'scripts/public-release-check.mjs')], {
  cwd: root,
  stdio: 'inherit'
});
if (check.status !== 0) process.exit(check.status ?? 1);

const output = execFileSync(
  'git',
  ['ls-files', '--cached', '--others', '--exclude-standard', '-z'],
  { cwd: root }
).toString('utf8');
const files = output.split('\0').filter(Boolean);
fs.mkdirSync(target, { recursive: true });
for (const file of files) {
  const source = path.join(root, file);
  const destination = path.join(target, file);
  fs.mkdirSync(path.dirname(destination), { recursive: true });
  fs.copyFileSync(source, destination);
  fs.chmodSync(destination, fs.statSync(source).mode);
}
console.log(`Exported ${files.length} sanitized files to ${target}.`);
console.log('Create a new Git repository in that directory only after the --release gate passes.');
