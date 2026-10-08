#!/usr/bin/env node
import { existsSync, writeFileSync } from 'node:fs';
import { createInterface } from 'node:readline/promises';
import { fileURLToPath } from 'node:url';
import { parseArgs } from 'node:util';
import { createQuickstartEnv, normalizePublicBaseUrl } from './onboarding.mjs';

const help = `Create a single-app .env without overwriting an existing file.
Usage: npm run setup
   or: npm run setup -- --tenant-id GUID --client-id GUID --public-url http://localhost:3000
No client secret is accepted on the command line. Edit .env after setup.
See docs/QUICKSTART.md for Entra consent and role assignment.`;
let rl;
try {
  const { values } = parseArgs({ options: {
    'tenant-id': { type: 'string' }, 'client-id': { type: 'string' },
    'public-url': { type: 'string' }, help: { type: 'boolean', short: 'h' }
  } });
  if (values.help) console.log(help);
  else {
    const target = fileURLToPath(new URL('../.env', import.meta.url));
    if (existsSync(target)) throw new Error('EXISTS');
    let tenantId = values['tenant-id'];
    let clientId = values['client-id'];
    let publicBaseUrl = values['public-url'];
    if ((!tenantId || !clientId) && !(process.stdin.isTTY && process.stdout.isTTY)) throw new Error('INPUT');
    if (process.stdin.isTTY && process.stdout.isTTY) {
      rl = createInterface({ input: process.stdin, output: process.stdout });
      tenantId ||= (await rl.question('21V Tenant ID: ')).trim();
      clientId ||= (await rl.question('MCP API Application (client) ID: ')).trim();
      publicBaseUrl ||= (await rl.question('Public base URL [http://localhost:3000]: ')).trim();
    }
    publicBaseUrl = normalizePublicBaseUrl(publicBaseUrl || 'http://localhost:3000');
    writeFileSync(target, createQuickstartEnv({ tenantId, clientId, publicBaseUrl }), { flag: 'wx', mode: 0o600 });
    console.log('Created .env. Existing files are never overwritten.');
    console.log('1. Edit MS_CLIENT_SECRET: use the secret VALUE, not its ID.');
    console.log('2. Complete Entra consent and assign mcp.users to the test user.');
    console.log('3. Run npm run doctor, npm ci, npm run build, then npm run start:http.');
    console.log(`Register this exact Web callback: ${publicBaseUrl}/oauth/microsoft/callback`);
    console.log('Only the profile module is loaded; authentication, App Roles and confirmation remain enabled.');
  }
} catch (error) {
  if (error.code === 'EEXIST' || error.message === 'EXISTS') console.error('Setup refused: .env already exists. Edit it manually; no file was changed.');
  else console.error('Setup failed. Check GUIDs, public URL, required arguments and directory access. Use npm run setup -- --help.');
  process.exitCode = 1;
} finally { rl?.close(); }
