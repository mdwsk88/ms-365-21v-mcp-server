import assert from 'node:assert/strict';
import test from 'node:test';

import {
  graphScopeEnabled,
  requestableGraphScope,
  requireGraphScope,
  scopeCovers
} from '../dist/scope-policy.js';

test('scope coverage follows the Read, ReadWrite, All and Shared hierarchy', () => {
  assert.equal(scopeCovers('Mail.Read', 'mail.read'), true);
  assert.equal(scopeCovers('https://microsoftgraph.chinacloudapi.cn/Mail.ReadWrite', 'Mail.Read'), true);
  assert.equal(scopeCovers('Files.ReadWrite.All', 'Files.Read'), true);
  assert.equal(scopeCovers('Files.ReadWrite.All', 'Files.ReadWrite'), true);
  assert.equal(scopeCovers('Files.Read.All', 'Files.Read'), true);
  assert.equal(scopeCovers('Mail.ReadWrite.Shared', 'Mail.Read.Shared'), true);
  assert.equal(scopeCovers('Sites.Manage.All', 'Sites.Read.All'), true);
  assert.equal(scopeCovers('User.Read.All', 'User.ReadBasic.All'), true);
  assert.equal(scopeCovers('Chat.ReadWrite', 'Chat.Read'), true);

  assert.equal(scopeCovers('Files.Read', 'Files.Read.All'), false);
  assert.equal(scopeCovers('Files.Read.All', 'Files.ReadWrite'), false);
  assert.equal(scopeCovers('Mail.Read', 'Mail.Read.Shared'), false);
  assert.equal(scopeCovers('Mail.ReadWrite', 'Mail.Send'), false);
  assert.equal(scopeCovers('Calendars.Read', 'Contacts.Read'), false);
});

test('deny-list wins and an allow-list fails closed', () => {
  const open = { disabledGraphScopes: [], allowedGraphScopes: [] };
  assert.equal(graphScopeEnabled(open, 'ChannelMessage.Read.All'), true);

  const denied = { disabledGraphScopes: ['ChannelMessage.Read.All'], allowedGraphScopes: ['ChannelMessage.Read.All'] };
  assert.equal(graphScopeEnabled(denied, 'channelmessage.read.all'), false);

  const allowed = { disabledGraphScopes: [], allowedGraphScopes: ['Mail.ReadWrite', 'Files.Read.All'] };
  assert.equal(graphScopeEnabled(allowed, 'Mail.Read'), true);
  assert.equal(graphScopeEnabled(allowed, 'Files.Read'), true);
  assert.equal(graphScopeEnabled(allowed, 'Mail.Read.Shared'), false);
  assert.throws(
    () => requireGraphScope(allowed, 'Mail.Read.Shared', 'Reading a shared mailbox'),
    error => error.code === 'scope_not_enabled' && /Mail\.Read\.Shared/.test(error.message)
  );
});

test('OBO requests ask for the granted covering scope when only a broader one is allowed', () => {
  assert.equal(requestableGraphScope({ allowedGraphScopes: [] }, 'Mail.Read'), 'Mail.Read');
  assert.equal(requestableGraphScope({ allowedGraphScopes: ['mail.read', 'mail.readwrite'] }, 'Mail.Read'), 'Mail.Read');
  assert.equal(requestableGraphScope({ allowedGraphScopes: ['mail.readwrite'] }, 'Mail.Read'), 'mail.readwrite');
  assert.equal(requestableGraphScope({ allowedGraphScopes: ['files.read.all'] }, 'Files.ReadWrite'), 'Files.ReadWrite');
});
