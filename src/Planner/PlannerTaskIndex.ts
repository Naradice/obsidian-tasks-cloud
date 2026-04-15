import type { App } from 'obsidian';

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

/**
 * Snapshot of the Obsidian task state at the time it was last pushed to Planner.
 * Used to detect changes without a full Graph API round-trip on every cache update.
 */
export interface SyncSnapshot {
    title: string;
    dueDate: string | null; // YYYY-MM-DD or null
    priority: string; // Priority enum value
    statusSymbol: string; // e.g. ' ', '/', 'x', '-'
    bucketId: string;
}

export interface IndexEntry {
    plannerId: string;
    planId: string;
    bucketId: string;
    etag: string;
    lastSyncedAt: number; // Unix ms
    snapshot: SyncSnapshot;
}

interface IndexData {
    version: number;
    /** Key = Tasks plugin id (the 🆔 field value) */
    tasks: Record<string, IndexEntry>;
}

// ---------------------------------------------------------------------------
// PlannerTaskIndex
// ---------------------------------------------------------------------------

/**
 * Persists the mapping between Tasks-plugin IDs and Planner task IDs.
 * Stored at  <vault>/.obsidian/plugins/<pluginId>/planner-index.json
 */
export class PlannerTaskIndex {
    private readonly app: App;
    private readonly indexPath: string;
    private data: IndexData = { version: 1, tasks: {} };

    constructor(app: App, pluginId: string) {
        this.app = app;
        this.indexPath = `${app.vault.configDir}/plugins/${pluginId}/planner-index.json`;
    }

    async load(): Promise<void> {
        try {
            const raw = await this.app.vault.adapter.read(this.indexPath);
            this.data = JSON.parse(raw) as IndexData;
        } catch {
            // File doesn't exist yet — start fresh
            this.data = { version: 1, tasks: {} };
        }
    }

    async save(): Promise<void> {
        // Ensure the parent directory exists (it may not on first run or when the
        // plugin folder name differs from the manifest id).
        const dir = this.indexPath.substring(0, this.indexPath.lastIndexOf('/'));
        if (!(await this.app.vault.adapter.exists(dir))) {
            await this.app.vault.adapter.mkdir(dir);
        }
        await this.app.vault.adapter.write(this.indexPath, JSON.stringify(this.data, null, 2));
    }

    // -----------------------------------------------------------------------
    // Query
    // -----------------------------------------------------------------------

    isLinked(tasksId: string): boolean {
        return tasksId !== '' && tasksId in this.data.tasks;
    }

    getByTasksId(tasksId: string): IndexEntry | undefined {
        return this.data.tasks[tasksId];
    }

    getByPlannerId(plannerId: string): { tasksId: string; entry: IndexEntry } | undefined {
        for (const [tasksId, entry] of Object.entries(this.data.tasks)) {
            if (entry.plannerId === plannerId) return { tasksId, entry };
        }
        return undefined;
    }

    getAllTasksIds(): string[] {
        return Object.keys(this.data.tasks);
    }

    // -----------------------------------------------------------------------
    // Mutations
    // -----------------------------------------------------------------------

    upsert(tasksId: string, entry: IndexEntry): void {
        this.data.tasks[tasksId] = entry;
    }

    updateEtagAndSnapshot(tasksId: string, etag: string, snapshot: SyncSnapshot): void {
        const entry = this.data.tasks[tasksId];
        if (entry) {
            entry.etag = etag;
            entry.snapshot = snapshot;
            entry.lastSyncedAt = Date.now();
        }
    }

    remove(tasksId: string): void {
        delete this.data.tasks[tasksId];
    }
}
