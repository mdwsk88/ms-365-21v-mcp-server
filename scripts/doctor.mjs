#!/usr/bin/env node
import { fileURLToPath } from 'node:url';
import { parseArgs } from 'node:util';
import { diagnoseEnv, readEnvironment } from './onboarding.mjs';

const jsonRequested = process.argv.includes('--json');
try {
  const { values } = parseArgs({ options: { json: { type: 'boolean' }, help: { type: 'boolean', short: 'h' }, 'config': { type: 'string' } } });
  if (values.help) console.log('Usage: npm run doctor -- [--json] [--config PATH]\nOffline HTTP/OAuth configuration checks. No network calls, cloud changes or credential values in output.\nDefault: repository .env; process environment values take precedence.');
  else {
    const file = values['config'] || fileURLToPath(new URL('../.env', import.meta.url));
    const { env, fileFound } = readEnvironment(file);
    const report = diagnoseEnv(env);
    if (!fileFound) {
      report.checks.push({ level: 'warning', code: 'ENV_FILE_MISSING', message: 'No env file found; checking the process environment only. Use npm run setup for a starter.' });
      report.warnings += 1;
    }
    if (values.json) console.log(JSON.stringify(report, null, 2));
    else {
      console.log(`MS 365-21V preflight: ${report.errors} error(s), ${report.warnings} warning(s).`);
      for (const check of report.checks) console.log(`[${check.level.toUpperCase()}] ${check.code}: ${check.message}`);
      console.log(report.ok ? 'Offline preflight passed. Next: build, start, connect and test auth_status / graph_get_me.' : 'Fix the errors above before connecting. See docs/QUICKSTART.md.');
      console.log(report.note);
    }
    process.exitCode = report.ok ? 0 : 1;
  }
} catch {
  const failure = { ok: false, errors: 1, warnings: 0, checks: [{ level: 'error', code: 'DOCTOR_INPUT', message: 'Unable to read configuration or parse arguments. Check file access and use --help.' }] };
  if (jsonRequested) console.log(JSON.stringify(failure, null, 2));
  else console.error(failure.checks[0].message);
  process.exitCode = 1;
}
