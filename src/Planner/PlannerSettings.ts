/**
 * Planner priority integers as defined by the Microsoft Graph API.
 *   0 = Urgent, 1 = Important, 5 = Medium, 9 = Low
 */
export type PlannerPriority = 0 | 1 | 5 | 9;

/** Maps an Obsidian tag (e.g. "#work") to a specific Planner plan+bucket. */
export interface TagBucketMapping {
    tag: string;
    planId: string;
    planTitle: string;
    bucketId: string;
    bucketTitle: string;
}

export interface PlannerSettings {
    enabled: boolean;

    // Azure AD app registration
    tenantId: string;
    clientId: string;

    // OAuth tokens — stored in plugin data, never displayed in plain text
    accessToken: string;
    accessTokenExpiresAt: number; // Unix timestamp ms
    refreshToken: string;

    // Authenticated user info (fetched after login)
    userId: string;
    userDisplayName: string;

    // Default target for tasks with no matching tag (usually "Private Tasks" plan)
    defaultPlanId: string;
    defaultPlanTitle: string;
    defaultBucketId: string;
    defaultBucketTitle: string;

    /**
     * Additional buckets (within defaultPlanId) to watch for pull-sync and import.
     * The defaultBucketId is always implicitly watched; these are extras.
     * Empty = only the default bucket is watched.
     */
    watchedBucketIds: string[];

    // Per-tag routing: first matching tag wins
    tagMappings: TagBucketMapping[];

    // Obsidian Priority → Planner priority number
    priorityMapping: {
        highest: PlannerPriority;
        high: PlannerPriority;
        medium: PlannerPriority;
        none: PlannerPriority;
        low: PlannerPriority;
        lowest: PlannerPriority;
    };

    // Automatically add a 🆔 ID to new tasks so they can be stably tracked
    autoAssignTaskIds: boolean;

    // Pull sync triggers
    syncOnFileOpen: boolean;
    syncIntervalMinutes: number; // 0 = disabled
}

/**
 * Returns all plan IDs that should be watched: the default plan plus any
 * plans referenced in tag mappings.
 */
export function allWatchedPlanIds(ps: PlannerSettings): string[] {
    const ids = new Set<string>();
    if (ps.defaultPlanId) ids.add(ps.defaultPlanId);
    for (const m of ps.tagMappings) {
        if (m.planId) ids.add(m.planId);
    }
    return Array.from(ids);
}

/**
 * Returns all bucket IDs to watch: the default bucket, any extras, plus
 * buckets from tag mappings (which may span multiple plans).
 * Deduplicates automatically.
 */
export function allWatchedBucketIds(ps: PlannerSettings): string[] {
    const ids = new Set<string>();
    if (ps.defaultBucketId) ids.add(ps.defaultBucketId);
    for (const id of ps.watchedBucketIds) {
        if (id) ids.add(id);
    }
    for (const m of ps.tagMappings) {
        if (m.bucketId) ids.add(m.bucketId);
    }
    return Array.from(ids);
}

export const DEFAULT_PLANNER_SETTINGS: PlannerSettings = {
    enabled: false,
    tenantId: '',
    clientId: '',
    accessToken: '',
    accessTokenExpiresAt: 0,
    refreshToken: '',
    userId: '',
    userDisplayName: '',
    defaultPlanId: '',
    defaultPlanTitle: '',
    defaultBucketId: '',
    defaultBucketTitle: '',
    watchedBucketIds: [],
    tagMappings: [],
    priorityMapping: {
        highest: 0, // Urgent
        high: 1, // Important
        medium: 5, // Medium
        none: 5, // Medium
        low: 9, // Low
        lowest: 9, // Low
    },
    autoAssignTaskIds: true,
    syncOnFileOpen: true,
    syncIntervalMinutes: 5,
};
