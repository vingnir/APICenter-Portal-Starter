/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation.
 *  Licensed under the MIT License.
 *--------------------------------------------------------------------------------------------*/

import { authService, configService } from "../util/useHttpClient";
import { IHttpClient, Method } from "./IHttpClient";

function normalizeHost(raw: string): string {
    // Strip scheme if present and any path that snuck in
    const noScheme = raw.replace(/^https?:\/\//i, "").split("/")[0];
    // Force data-plane host (defensive if someone typed the portal host)
    return noScheme.replace(".portal.", ".data.");
}

function joinUrl(base: string, path: string): string {
    return `${base.replace(/\/+$/, "")}/${path.replace(/^\/+/, "")}`;
}

export class HttpClient implements IHttpClient {
    public async fetchData(url: string, method: Method = Method.GET): Promise<any> {
        const accessToken = await authService.getAccessToken();
        const settings = await configService.getSettings();

        // ---- Config expectations -------------------------------------------------
        // settings.dataApiHostName MUST be host-only, e.g.
        // "meab-api-portal.data.swedencentral.azure-apicenter.ms"
        const host = normalizeHost(settings.dataApiHostName);

        // workspace selection (first configured or 'default')
        const workspace =
            (Array.isArray(settings.workspaces) && settings.workspaces[0]) || settings.workspace || "default";

        // If caller didn't already include /workspaces/<name>, add it.
        // Example target: https://<host>/workspaces/<ws>/<url>
        const hasWorkspacePrefix = /^\/?workspaces\//i.test(url);
        const pathWithWorkspace = hasWorkspacePrefix ? url : joinUrl(`/workspaces/${workspace}`, url);

        const requestUrl = `https://${host}${pathWithWorkspace.startsWith("/") ? "" : "/"}${pathWithWorkspace}`;

        const headers: HeadersInit = {
            Accept: "application/json",
            "Content-Type": "application/json",
        };
        if (accessToken) headers.Authorization = `Bearer ${accessToken}`;

        const response = await fetch(requestUrl, { method, headers });

        // Handle common auth/permission issues gracefully for the UI
        if (response.status === 401 || response.status === 403) {
            if (accessToken) localStorage.setItem("MS_APIC_DEVPORTAL_isRestricted", "true");
            return null;
        }
        if (response.status === 404) return null;

        // If CORS blocks the response, this will throw; helpful console for debugging
        try {
            return await response.json();
        } catch (e) {
            console.error("Failed to parse API Center response:", requestUrl, e);
            throw e;
        }
    }
}
