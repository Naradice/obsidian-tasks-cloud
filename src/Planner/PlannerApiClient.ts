import { requestUrl, type RequestUrlParam } from 'obsidian';
import { PlannerAuth } from './PlannerAuth';
import type { PlannerSettings } from './PlannerSettings';

// ---------------------------------------------------------------------------
// Graph API response shapes (minimal — only fields we use)
// ---------------------------------------------------------------------------

export interface PlannerTask {
    id: string;
    planId: string;
    bucketId: string;
    title: string;
    /** 0 = not started, 50 = in progress, 100 = completed */
    percentComplete: number;
    /** 0 = Urgent, 1 = Important, 5 = Medium, 9 = Low */
    priority: number;
    /** ISO 8601 UTC, or null */
    dueDateTime: string | null;
    createdDateTime: string;
    completedDateTime: string | null;
    assignments: Record<string, { '@odata.type': string; orderHint: string }>;
}

export interface PlannerBucket {
    id: string;
    planId: string;
    name: string;
    orderHint: string;
}

export interface PlannerPlan {
    id: string;
    title: string;
    owner: string;
    createdDateTime: string;
}

export interface GraphUser {
    id: string;
    displayName: string;
    mail: string;
}

interface GraphListResponse<T> {
    value: T[];
    '@odata.nextLink'?: string;
}

// ---------------------------------------------------------------------------
// Client
// ---------------------------------------------------------------------------

export class PlannerApiClient {
    private static readonly BASE = 'https://graph.microsoft.com/v1.0';

    private readonly getSettings: () => PlannerSettings;
    private readonly saveSettings: (patch: Partial<PlannerSettings>) => Promise<void>;

    constructor(
        getSettings: () => PlannerSettings,
        saveSettings: (patch: Partial<PlannerSettings>) => Promise<void>,
    ) {
        this.getSettings = getSettings;
        this.saveSettings = saveSettings;
    }

    // -----------------------------------------------------------------------
    // Token management
    // -----------------------------------------------------------------------

    private async getAccessToken(): Promise<string> {
        const s = this.getSettings();
        if (!PlannerAuth.isTokenExpired(s)) return s.accessToken;

        const tokens = await PlannerAuth.refreshAccessToken(s.tenantId, s.clientId, s.refreshToken);
        const patch: Partial<PlannerSettings> = {
            accessToken: tokens.access_token,
            accessTokenExpiresAt: Date.now() + tokens.expires_in * 1000,
        };
        if (tokens.refresh_token) patch.refreshToken = tokens.refresh_token;
        await this.saveSettings(patch);
        return tokens.access_token;
    }

    // -----------------------------------------------------------------------
    // Generic request helper
    // -----------------------------------------------------------------------

    private async request<T>(
        method: string,
        path: string,
        body?: unknown,
        extraHeaders?: Record<string, string>,
    ): Promise<{ data: T; etag: string | null; status: number }> {
        const token = await this.getAccessToken();

        const params: RequestUrlParam = {
            url: `${PlannerApiClient.BASE}${path}`,
            method,
            headers: {
                Authorization: `Bearer ${token}`,
                'Content-Type': 'application/json',
                ...extraHeaders,
            },
            throw: false,
        };

        if (body !== undefined) params.body = JSON.stringify(body);

        const response = await requestUrl(params);

        if (response.status === 401) throw new Error('Planner: unauthorized — please re-authenticate in settings.');
        if (response.status === 403) throw new Error('Planner: access denied — check your Azure app permissions.');
        if (response.status >= 400) {
            const body = (response.json as { error?: { message?: string } }) ?? {};
            throw new Error(`Planner API ${response.status}: ${body.error?.message ?? JSON.stringify(body)}`);
        }

        const etag = (response.headers as Record<string, string>)?.['etag'] ?? null;
        const data = response.status === 204 ? (null as unknown as T) : (response.json as T);
        return { data, etag, status: response.status };
    }

    // -----------------------------------------------------------------------
    // User
    // -----------------------------------------------------------------------

    async getMe(): Promise<GraphUser> {
        const { data } = await this.request<GraphUser>('GET', '/me');
        return data;
    }

    // -----------------------------------------------------------------------
    // Plans & buckets (for settings UI dropdowns)
    // -----------------------------------------------------------------------

    async getMyPlans(): Promise<PlannerPlan[]> {
        const { data } = await this.request<GraphListResponse<PlannerPlan>>('GET', '/me/planner/plans');
        return data.value;
    }

    async getPlanBuckets(planId: string): Promise<PlannerBucket[]> {
        const { data } = await this.request<GraphListResponse<PlannerBucket>>(
            'GET',
            `/planner/plans/${planId}/buckets`,
        );
        return data.value;
    }

    // -----------------------------------------------------------------------
    // Tasks
    // -----------------------------------------------------------------------

    async getMyTasks(): Promise<PlannerTask[]> {
        const { data } = await this.request<GraphListResponse<PlannerTask>>('GET', '/me/planner/tasks');
        return data.value;
    }

    async getPlanTasks(planId: string): Promise<PlannerTask[]> {
        const { data } = await this.request<GraphListResponse<PlannerTask>>(
            'GET',
            `/planner/plans/${planId}/tasks`,
        );
        return data.value;
    }

    /** Fetch a single task and its current ETag. */
    async getTask(taskId: string): Promise<{ task: PlannerTask; etag: string }> {
        const { data, etag } = await this.request<PlannerTask>('GET', `/planner/tasks/${taskId}`);
        if (!etag) throw new Error(`No ETag returned for task ${taskId}`);
        return { task: data, etag };
    }

    /** Create a new task. Returns the created task and its ETag for future updates. */
    async createTask(
        planId: string,
        bucketId: string,
        title: string,
        dueDateTime: string | null,
        priority: number,
        percentComplete: number,
        userId: string,
    ): Promise<{ task: PlannerTask; etag: string }> {
        const body: Record<string, unknown> = { planId, bucketId, title, priority, percentComplete };
        if (dueDateTime) body.dueDateTime = dueDateTime;
        if (userId) {
            body.assignments = {
                [userId]: {
                    '@odata.type': '#microsoft.graph.plannerAssignment',
                    orderHint: ' !',
                },
            };
        }

        const { data, etag } = await this.request<PlannerTask>('POST', '/planner/tasks', body);
        if (!etag) throw new Error('No ETag returned from task creation');
        return { task: data, etag };
    }

    /**
     * Patch specific fields on an existing task.
     * The ETag from the last GET/POST/PATCH must be supplied for optimistic concurrency.
     * Returns the new ETag.
     */
    async updateTask(
        taskId: string,
        etag: string,
        changes: Partial<Pick<PlannerTask, 'title' | 'dueDateTime' | 'priority' | 'percentComplete' | 'bucketId'>>,
    ): Promise<string> {
        const { etag: newEtag } = await this.request<null>('PATCH', `/planner/tasks/${taskId}`, changes, {
            'If-Match': etag,
        });
        // Planner returns 204 No Content on successful PATCH, so we keep the old ETag as fallback
        return newEtag ?? etag;
    }

    /** Delete a task. Requires the current ETag. */
    async deleteTask(taskId: string, etag: string): Promise<void> {
        await this.request<null>('DELETE', `/planner/tasks/${taskId}`, undefined, { 'If-Match': etag });
    }
}
