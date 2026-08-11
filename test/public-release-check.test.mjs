import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { spawnSync } from 'node:child_process';
import test from 'node:test';

const sourceScript = new URL('../scripts/public-release-check.mjs', import.meta.url);

function runCheck(directory, ...args) {
  return spawnSync(process.execPath, ['scripts/public-release-check.mjs', ...args], {
    cwd: directory,
    encoding: 'utf8'
  });
}

test('public release checker works before git init and enforces the license gate', async (t) => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'public-release-check-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  await fs.mkdir(path.join(directory, 'scripts'));
  await fs.copyFile(sourceScript, path.join(directory, 'scripts/public-release-check.mjs'));
  await fs.writeFile(path.join(directory, 'README.md'), '# Sanitized project\n');
  await fs.writeFile(path.join(directory, 'SECURITY.md'), '# Security\n');
  await fs.writeFile(path.join(directory, 'CONTRIBUTING.md'), '# Contributing\n');

  const preparation = runCheck(directory);
  assert.equal(preparation.status, 0, preparation.stderr);
  assert.match(preparation.stdout, /preparation check passed/);

  const blockedRelease = runCheck(directory, '--release');
  assert.equal(blockedRelease.status, 1);
  assert.match(blockedRelease.stderr, /LICENSE/);

  await fs.writeFile(path.join(directory, 'LICENSE'), 'Approved license text\n');
  const approvedRelease = runCheck(directory, '--release');
  assert.equal(approvedRelease.status, 0, approvedRelease.stderr);

  await fs.writeFile(
    path.join(directory, '.env.example'),
    'MS_CLIENT_ID=11111111-1111-1111-1111-111111111111\n'
  );
  const leakedIdentifier = runCheck(directory);
  assert.equal(leakedIdentifier.status, 1);
  assert.match(leakedIdentifier.stderr, /real tenant or application UUID/);
});
