/**
 * Microsoft Graph API client factory.
 * Returns an authenticated Client instance ready for API calls.
 */

import { Client } from "@microsoft/microsoft-graph-client";
import { getAccessToken } from "./auth";

/**
 * Creates and returns an authenticated Microsoft Graph client.
 * The auth provider calls `getAccessToken()` on every request, which
 * handles caching and automatic token refresh transparently.
 */
export async function getGraphClient(): Promise<Client> {
  return Client.init({
    authProvider: async (done) => {
      try {
        const token = await getAccessToken();
        done(null, token);
      } catch (error) {
        done(error as Error, null);
      }
    },
  });
}
