import type { App } from 'obsidian';

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

export type QueuedOperationType = 'create' | 'update' | 'delete';

export interface QueuedOperation {
    /** Random ID used for deduplication and removal. */
    id: string;
    type: QueuedOperationType;
    /** Tasks-plugin ID of the Obsidian task this operation targets. */
    tasksId: string;
    timestamp: number;
    // create / update
    payload?: Record<string, unknown>;
    // update / delete
    plannerId?: string;
    etag?: string;
}

interface QueueData {
    version: number;
    operations: QueuedOperation[];
}

// ---------------------------------------------------------------------------
// PlannerOfflineQueue
// ---------------------------------------------------------------------------

/**
 * Persists pending Planner operations to disk so they survive an Obsidian restart.
 *
 * Collapse rules on enqueue (prevents redundant API calls):
 *   - A new `delete` for tasksId X removes all existing ops for X and takes over.
 *   - A new `update` for tasksId X merges its payload into any existing pending update.
 *   - A new `create` is always appended (duplicates can't arise once IDs are tracked).
 */
export class PlannerOfflineQueue {
    private readonly app: App;
    private readonly queuePath: string;
    private data: QueueData = { version: 1, operations: [] };

    constructor(app: App, pluginId: string) {
        this.app = app;
        this.queuePath = `${app.vault.configDir}/plugins/${pluginId}/planner-queue.json`;
    }

    async load(): Promise<void> {
        try {
            const raw = await this.app.vault.adapter.read(this.queuePath);
            this.data = JSON.parse(raw) as QueueData;
        } catch {
            this.data = { version: 1, operations: [] };
        }
    }

    async save(): Promise<void> {
        const dir = this.queuePath.substring(0, this.queuePath.lastIndexOf('/'));
        if (!(await this.app.vault.adapter.exists(dir))) {
            await this.app.vault.adapter.mkdir(dir);
        }
        await this.app.vault.adapter.write(this.queuePath, JSON.stringify(this.data, null, 2));
    }

    // -----------------------------------------------------------------------
    // Enqueue (with collapse)
    // -----------------------------------------------------------------------

    enqueue(op: Omit<QueuedOperation, 'id' | 'timestamp'>): void {
        if (op.type === 'delete') {
            // Delete supersedes any pending create or update for this task
            this.data.operations = this.data.operations.filter((o) => o.tasksId !== op.tasksId);
        } else if (op.type === 'update') {
            const existingIdx = this.data.operations.findIndex(
                (o) => o.tasksId === op.tasksId && o.type === 'update',
            );
            if (existingIdx !== -1) {
                // Merge payloads — newer fields win
                this.data.operations[existingIdx].payload = {
                    ...this.data.operations[existingIdx].payload,
                    ...op.payload,
                };
                // Also refresh etag if provided
                if (op.etag) this.data.operations[existingIdx].etag = op.etag;
                return;
            }
        }

        this.data.operations.push({
            id: Math.random().toString(36).slice(2, 10),
            timestamp: Date.now(),
            ...op,
        });
    }

    // -----------------------------------------------------------------------
    // Dequeue / inspect
    // -----------------------------------------------------------------------

    peek(): QueuedOperation | undefined {
        return this.data.operations[0];
    }

    dequeue(): void {
        this.data.operations.shift();
    }

    removeById(id: string): void {
        this.data.operations = this.data.operations.filter((o) => o.id !== id);
    }

    get length(): number {
        return this.data.operations.length;
    }

    isEmpty(): boolean {
        return this.data.operations.length === 0;
    }
}
