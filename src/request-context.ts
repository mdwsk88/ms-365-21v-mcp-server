import { AsyncLocalStorage } from 'node:async_hooks';
import type { JWTPayload } from 'jose';

export type RequestContext = {
  requestId: string;
  userAssertion?: string;
  inboundClaims?: JWTPayload;
};

const storage = new AsyncLocalStorage<RequestContext>();

export function runWithRequestContext<T>(context: RequestContext, handler: () => T): T {
  return storage.run(context, handler);
}

export function getRequestContext(): RequestContext | undefined {
  return storage.getStore();
}
