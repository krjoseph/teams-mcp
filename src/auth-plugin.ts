import { useIdentityPlugin } from "@azure/identity";
import { cachePersistencePlugin } from "@azure/identity-cache-persistence";

// Enable persistent token caching (stores refresh tokens)
// This must only be called once globally
let pluginInitialized = false;

export function initializeCachePlugin(): void {
  if (!pluginInitialized) {
    useIdentityPlugin(cachePersistencePlugin);
    pluginInitialized = true;
  }
}
