import { z } from 'zod/v4';
import {
  clearAuth,
  completeDeviceLogin,
  getPendingDeviceLogin,
  getTokenCache,
  secondsToIso,
  startDeviceLogin
} from '../oauth.js';
import { getRequestContext } from '../request-context.js';
import { describeTool, runTool } from './results.js';
import type { ToolModule } from './types.js';

export const authModule: ToolModule = {
  category: 'gateway',
  displayName: 'Gateway Authentication',
  description: 'Authentication diagnostics and local development helpers.',
  alwaysEnabled: true,
  toolNames: ['auth_start_device_login', 'auth_complete_device_login', 'auth_status', 'auth_clear_local_cache'],
  activeToolNames(config) {
    return config.transport === 'stdio'
      ? ['auth_start_device_login', 'auth_complete_device_login', 'auth_status', 'auth_clear_local_cache']
      : ['auth_status'];
  },
  register(server, config) {
    if (config.transport === 'stdio') {
      server.registerTool(
        'auth_start_device_login',
        {
          title: 'Start Local Device Login',
          description: describeTool(
            'Development only. Start a 21V Entra device-code login and cache the result locally.',
            ['本地登录', '设备代码登录', '开发调试登录']
          ),
          inputSchema: {
            scopes: z
              .array(z.string())
              .optional()
              .describe('Optional one-time scope override. Defaults to MS_GRAPH_SCOPES.')
          }
        },
        async ({ scopes }) =>
          runTool(async () => {
            const state = await startDeviceLogin(config, scopes?.length ? scopes : config.scopes);
            return {
              userCode: state.userCode,
              verificationUri: state.verificationUri,
              verificationUriComplete: state.verificationUriComplete,
              expiresAt: secondsToIso(state.expiresAt),
              intervalSeconds: state.intervalSeconds,
              message: state.message,
              scopes: state.scopes
            };
          })
      );

      server.registerTool(
        'auth_complete_device_login',
        {
          title: 'Complete Local Device Login',
          description: describeTool(
            'Development only. Poll the token endpoint after the user enters the device code, then cache the token locally. Token values are not returned.',
            ['完成本地登录', '完成设备码登录', '保存本地令牌']
          ),
          inputSchema: {
            timeoutSeconds: z
              .number()
              .int()
              .min(0)
              .max(900)
              .optional()
              .describe('Polling timeout in seconds. Defaults to 120.')
          }
        },
        async ({ timeoutSeconds }) =>
          runTool(async () => {
            const token = await completeDeviceLogin(config, timeoutSeconds ?? 120);
            return {
              authenticated: true,
              tokenType: token.tokenType,
              scopes: token.scopes,
              expiresAt: secondsToIso(token.expiresAt),
              hasRefreshToken: Boolean(token.refreshToken)
            };
          })
      );
    }

    server.registerTool(
      'auth_status',
      {
        title: 'Check Auth Status',
        description: describeTool(
          'Inspect MCP auth state, local token state, remote OAuth/OBO configuration, and current request user details. Token values are never returned.',
          ['查看认证状态', '检查登录状态', '诊断授权问题']
        )
      },
      async () =>
        runTool(async () => {
          const token = await getTokenCache(config);
          const pending = await getPendingDeviceLogin(config);
          const requestContext = getRequestContext();
          return {
            transport: config.transport,
            localDeviceCodeConfigured: config.missing.length === 0,
            localDeviceCodeMissing: config.missing,
            remoteConfigured: config.missingRemote.length === 0,
            remoteMissing: config.missingRemote,
            resourceUrl: config.resourceUrl,
            resourceMetadataUrl: config.resourceMetadataUrl,
            authorizationServers: config.authorizationServers,
            authorizationScopes: config.authorizationScopes,
            requiredTokenScopes: config.requiredTokenScopes,
            authorityHost: config.authorityHost,
            oidcMetadataUrl: config.oidcMetadataUrl,
            graphBaseUrl: config.graphBaseUrl,
            localScopes: config.scopes,
            oboScopes: config.oboScopes,
            tokenCachePath: config.tokenCachePath,
            authenticatedLocally: Boolean(token),
            tokenExpiresAt: secondsToIso(token?.expiresAt),
            hasRefreshToken: Boolean(token?.refreshToken),
            requestHasUserAssertion: Boolean(requestContext?.userAssertion),
            requestSubject: requestContext?.inboundClaims?.sub,
            requestUser: requestContext?.inboundClaims?.preferred_username ?? requestContext?.inboundClaims?.upn,
            pendingDeviceLogin: pending
              ? {
                  userCode: pending.userCode,
                  verificationUri: pending.verificationUri,
                  verificationUriComplete: pending.verificationUriComplete,
                  expiresAt: secondsToIso(pending.expiresAt),
                  scopes: pending.scopes
                }
              : undefined,
            rawGraphGetEnabled: config.enableRawGraphGet
          };
        })
    );

    if (config.transport === 'stdio') {
      server.registerTool(
        'auth_clear_local_cache',
        {
          title: 'Clear Local Auth Cache',
          description: describeTool(
            'Development only. Delete locally cached tokens and pending device-code login state.',
            ['清除本地登录', '退出本地登录', '删除缓存令牌']
          )
        },
        async () =>
          runTool(async () => {
            await clearAuth(config);
            return { authenticated: false, cleared: true };
          })
      );
    }
  }
};
