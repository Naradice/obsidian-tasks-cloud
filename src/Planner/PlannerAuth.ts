import { requestUrl } from 'obsidian';
import type { PlannerSettings } from './PlannerSettings';

interface DeviceCodeResponse {
    device_code: string;
    user_code: string;
    verification_uri: string;
    expires_in: number;
    interval: number;
    message: string;
}

export interface TokenResponse {
    access_token: string;
    refresh_token?: string;
    expires_in: number;
}

// Scopes required for reading/writing Planner tasks and resolving the current user
const SCOPE = 'Tasks.ReadWrite User.Read offline_access';

export class PlannerAuth {
    /**
     * Step 1 of the device code flow.
     * Returns the user-facing code and verification URL to display in the settings UI.
     */
    static async startDeviceCodeFlow(tenantId: string, clientId: string): Promise<DeviceCodeResponse> {
        const response = await requestUrl({
            url: `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/devicecode`,
            method: 'POST',
            headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
            body: new URLSearchParams({ client_id: clientId, scope: SCOPE }).toString(),
            throw: false,
        });

        if (response.status !== 200) {
            const err = (response.json as Record<string, string>) ?? {};
            throw new Error(err.error_description ?? `Device code request failed (HTTP ${response.status})`);
        }

        return response.json as DeviceCodeResponse;
    }

    /**
     * Step 2 of the device code flow: poll until the user authorises the app.
     * Pass an AbortSignal to allow the caller to cancel (e.g. when the settings
     * modal is closed).
     */
    static async pollForToken(
        tenantId: string,
        clientId: string,
        deviceCode: string,
        intervalSeconds: number,
        signal: AbortSignal,
    ): Promise<TokenResponse> {
        let pollMs = intervalSeconds * 1000;

        while (!signal.aborted) {
            await new Promise<void>((resolve) => {
                const timer = setTimeout(resolve, pollMs);
                signal.addEventListener('abort', () => {
                    clearTimeout(timer);
                    resolve();
                });
            });

            if (signal.aborted) break;

            const response = await requestUrl({
                url: `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
                method: 'POST',
                headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                body: new URLSearchParams({
                    grant_type: 'urn:ietf:params:oauth:grant-type:device_code',
                    client_id: clientId,
                    device_code: deviceCode,
                }).toString(),
                throw: false,
            });

            if (response.status === 200) {
                return response.json as TokenResponse;
            }

            const err = (response.json as Record<string, string>) ?? {};
            switch (err.error) {
                case 'authorization_pending':
                    continue;
                case 'slow_down':
                    pollMs += 5000;
                    continue;
                case 'expired_token':
                    throw new Error('Authentication timed out. Please start again.');
                case 'access_denied':
                    throw new Error('Authentication was declined by the user.');
                default:
                    throw new Error(err.error_description ?? `Unexpected auth error: ${err.error}`);
            }
        }

        throw new Error('Authentication cancelled.');
    }

    /** Silently refresh the access token using the stored refresh token. */
    static async refreshAccessToken(tenantId: string, clientId: string, refreshToken: string): Promise<TokenResponse> {
        const response = await requestUrl({
            url: `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
            method: 'POST',
            headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
            body: new URLSearchParams({
                grant_type: 'refresh_token',
                client_id: clientId,
                refresh_token: refreshToken,
                scope: SCOPE,
            }).toString(),
            throw: false,
        });

        if (response.status !== 200) {
            const err = (response.json as Record<string, string>) ?? {};
            throw new Error(err.error_description ?? `Token refresh failed (HTTP ${response.status})`);
        }

        return response.json as TokenResponse;
    }

    /** True when the stored access token is missing or within 5 minutes of expiry. */
    static isTokenExpired(settings: PlannerSettings): boolean {
        if (!settings.accessToken) return true;
        return Date.now() >= settings.accessTokenExpiresAt - 5 * 60 * 1000;
    }

    /** True when the user has completed the device code flow at least once. */
    static isAuthenticated(settings: PlannerSettings): boolean {
        return settings.accessToken !== '' && settings.refreshToken !== '';
    }
}
