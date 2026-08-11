import { z } from 'zod/v4';
import { browserConfirmationStore, confirmationUserKey } from '../browser-confirmation.js';
import { hasConfirmationPolicy } from '../confirmation.js';
import { getRequestContext } from '../request-context.js';
import { describeTool } from './results.js';
import type { ToolModule } from './types.js';

export const confirmModule: ToolModule = {
  category: 'gateway',
  displayName: 'Operation Confirmation',
  description: 'Continues a sensitive operation after explicit human approval.',
  alwaysEnabled: true,
  toolNames: ['confirm_execute'],
  isEnabled: config => hasConfirmationPolicy(config),
  register(server) {
    server.registerTool(
      'confirm_execute',
      {
        title: 'Execute Approved Operation',
        description: describeTool(
          'Continue a pending sensitive operation after the same signed-in user approved its web preview. This tool never grants approval itself. Call it only after the user says the confirmation page was approved; the token is user-bound, one-time, and short-lived.',
          ['继续执行我刚刚在网页确认的操作', '执行已批准的敏感操作']
        ),
        inputSchema: {
          confirmToken: z.string().min(20).max(200).describe('One-time token returned by the pending operation.')
        }
      },
      async ({ confirmToken }) =>
        browserConfirmationStore.execute(
          confirmToken,
          confirmationUserKey(getRequestContext()?.inboundClaims)
        )
    );
  }
};
