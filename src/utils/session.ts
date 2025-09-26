import { AsyncLocalStorage } from 'async_hooks';

export interface SessionContext {
  sessionId: string;
}

export const sessionStorage = new AsyncLocalStorage<SessionContext>();
