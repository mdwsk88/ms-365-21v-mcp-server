import assert from 'node:assert/strict';
import test from 'node:test';

import {
  hasConfirmationPolicy,
  operationRequiresConfirmation
} from '../dist/confirmation.js';

test('confirm mode always gates email and Teams sends', () => {
  const config = { sendMode: 'confirm', confirmOperations: [] };
  assert.equal(operationRequiresConfirmation(config, 'send_email'), true);
  assert.equal(operationRequiresConfirmation(config, 'send_teams_message'), true);
  assert.equal(hasConfirmationPolicy(config), true);
});

test('automatic mode bypasses only send confirmation even when the generic policy is all', () => {
  const config = { sendMode: 'automatic', confirmOperations: ['all'] };
  assert.equal(operationRequiresConfirmation(config, 'send_email'), false);
  assert.equal(operationRequiresConfirmation(config, 'send_teams_message'), false);
  assert.equal(operationRequiresConfirmation(config, 'delete_drive_item'), true);
  assert.equal(hasConfirmationPolicy(config), true);
});

test('automatic mode disables confirmation routes when only send operations were listed', () => {
  const config = {
    sendMode: 'automatic',
    confirmOperations: ['send_email', 'send_teams_message']
  };
  assert.equal(hasConfirmationPolicy(config), false);
});
