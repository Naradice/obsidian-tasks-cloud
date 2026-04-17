import { Notice, type TFile, normalizePath } from 'obsidian';
import type TasksPlugin from '../main';
import type { Task } from '../Task/Task';
import { Task as TaskClass } from '../Task/Task';
import { StatusRegistry } from '../Statuses/StatusRegistry';
import { StatusType } from '../Statuses/StatusConfiguration';
import { replaceTaskWithTasks } from '../Obsidian/File';
import { getSettings, updateSettings } from '../Config/Settings';
import type { TasksEvents } from '../Obsidian/TasksEvents';
import { generateUniqueId } from '../Task/TaskDependency';
import { PlannerAuth } from './PlannerAuth';
import { PlannerApiClient } from './PlannerApiClient';
import { PlannerTaskIndex } from './PlannerTaskIndex';
import { PlannerTaskMapper } from './PlannerTaskMapper';
import { PlannerOfflineQueue } from './PlannerOfflineQueue';
import type { PlannerSettings } from './PlannerSettings';
import { allWatchedBucketIds } from './PlannerSettings';

// ---------------------------------------------------------------------------
// PlannerSyncManager
// ---------------------------------------------------------------------------

/**
 * Orchestrates all Planner ↔ Obsidian sync operations.
 *
 * Lifecycle:
 *   1. Instantiate in main.ts onload(), passing the plugin and events.
 *   2. Call initialize() after the workspace is ready.
 *   3. Call unload() in main.ts onunload().
 *
 * Push (Obsidian → Planner):
 *   Triggered by onTasksChanged(), which is called from the cache-update event.
 *   The manager diffs the new task list against a snapshot keyed by Tasks-plugin ID.
 *   Tasks without a 🆔 ID are auto-assigned one (when autoAssignTaskIds is enabled)
 *   so they can be stably tracked across edits.
 *
 * Pull (Planner → Obsidian):
 *   Triggered on file-open and on a configurable interval.
 *   Planner is always master on conflict.
 */
export class PlannerSyncManager {
    private readonly plugin: TasksPlugin;
    private readonly events: TasksEvents;

    readonly index: PlannerTaskIndex;
    readonly queue: PlannerOfflineQueue;
    private client: PlannerApiClient;

    /** Snapshot of Tasks-plugin-ID → Task, updated after every successful diff. */
    private previousTasks = new Map<string, Task>();

    /**
     * Composite keys (path:line) of tasks currently having a 🆔 ID auto-assigned.
     * Prevents duplicate assignments when the cache fires while the write is in flight.
     */
    private readonly pendingIdAssignments = new Set<string>();

    /**
     * Maps compositeKey (path:line) → assigned task ID.
     * Set the moment autoAssignId picks an ID (before the async file write).
     * Lets processTasksChanged restore the same ID if the line is subsequently
     * overwritten without the ID (e.g. by the edit modal reading a stale editor
     * buffer before autoAssignId's vault write was reflected).
     * Cleared when the task with the confirmed ID is processed in newMap.
     */
    private readonly assignedIdAtKey = new Map<string, string>();

    /**
     * Pending links from UC2: when the user picks an existing Planner task in the
     * modal but the Obsidian task has no 🆔 ID yet, we store the intent here.
     * Key = trimmed task description; value = Planner task ID to link.
     * Consumed in onTasksChanged when a matching new task arrives.
     */
    private readonly pendingLinks = new Map<string, string>();

    /**
     * Set on the very first onTasksChanged call so we can distinguish
     * pre-existing vault tasks from tasks genuinely created after the plugin
     * was enabled.  Keyed by compositeKey (path:line).
     */
    private readonly preExistingTaskKeys = new Set<string>();

    /** False until the first onTasksChanged snapshot has been taken. */
    private initialSnapshotBuilt = false;

    /** True while processTasksChanged is running, to prevent concurrent execution. */
    private isProcessingTasksChanged = false;

    /**
     * If onTasksChanged is called while already processing, the latest snapshot
     * is stored here and processed once the current run finishes.
     */
    private pendingTasksSnapshot: Task[] | null = null;

    private syncIntervalId: number | null = null;

    constructor(plugin: TasksPlugin, events: TasksEvents) {
        this.plugin = plugin;
        this.events = events;
        this.index = new PlannerTaskIndex(plugin.app, plugin.manifest.id);
        this.queue = new PlannerOfflineQueue(plugin.app, plugin.manifest.id);
        this.client = new PlannerApiClient(
            () => getSettings().plannerSettings,
            async (patch) => {
                updateSettings({ plannerSettings: { ...getSettings().plannerSettings, ...patch } });
                await plugin.saveSettings();
            },
        );
    }

    // -----------------------------------------------------------------------
    // Lifecycle
    // -----------------------------------------------------------------------

    async initialize(): Promise<void> {
        await this.index.load();
        await this.queue.load();
        this.registerListeners();
        this.startSyncInterval();

        // Attempt to flush any operations queued from a previous session
        if (!this.queue.isEmpty()) {
            await this.flushQueue();
        }
    }

    unload(): void {
        if (this.syncIntervalId !== null) {
            window.clearInterval(this.syncIntervalId);
            this.syncIntervalId = null;
        }
    }

    // -----------------------------------------------------------------------
    // Configuration helpers
    // -----------------------------------------------------------------------

    private getSettings(): PlannerSettings {
        return getSettings().plannerSettings;
    }

    private isReady(): boolean {
        const s = this.getSettings();
        return s.enabled && PlannerAuth.isAuthenticated(s) && s.defaultPlanId !== '';
    }

    // -----------------------------------------------------------------------
    // Event listeners
    // -----------------------------------------------------------------------

    private registerListeners(): void {
        // Pull on file open
        const fileOpenRef = this.plugin.app.workspace.on('file-open', (file: TFile | null) => {
            if (!file || !this.isReady()) return;
            if (this.getSettings().syncOnFileOpen) {
                this.pullSyncFile(file).catch((err) => console.error('Planner pullSyncFile error:', err));
            }
        });
        this.plugin.registerEvent(fileOpenRef);
    }

    startSyncInterval(): void {
        this.stopSyncInterval();
        const minutes = this.getSettings().syncIntervalMinutes;
        if (minutes <= 0) return;

        this.syncIntervalId = window.setInterval(
            () => {
                if (!this.isReady()) return;
                const file = this.plugin.app.workspace.getActiveFile();
                if (file) this.pullSyncFile(file).catch((err) => console.error('Planner interval pull error:', err));
            },
            minutes * 60 * 1000,
        );
    }

    stopSyncInterval(): void {
        if (this.syncIntervalId !== null) {
            window.clearInterval(this.syncIntervalId);
            this.syncIntervalId = null;
        }
    }

    // -----------------------------------------------------------------------
    // Push sync  (Obsidian → Planner)
    // -----------------------------------------------------------------------

    /**
     * Called by main.ts whenever the task cache emits an update.
     * Serialises execution: if a previous call is still in-flight the incoming
     * snapshot is buffered and processed once the current run finishes.
     */
    async onTasksChanged(allTasks: Task[]): Promise<void> {
        // Always update snapshot when not ready, so we never accumulate a
        // backlog of "new" tasks the first time the integration is enabled.
        if (!this.isReady()) {
            this.rebuildSnapshot(allTasks);
            return;
        }

        // ── First call: record every existing task location so we can
        //   distinguish pre-existing tasks from genuinely new ones.
        //   Do NOT auto-assign IDs or push anything here.
        if (!this.initialSnapshotBuilt) {
            this.initialSnapshotBuilt = true;
            for (const task of allTasks) {
                this.preExistingTaskKeys.add(this.compositeKey(task));
            }
            this.rebuildSnapshot(allTasks);
            return;
        }

        // If already processing, buffer the latest snapshot and return.
        // The running call will drain the buffer when it finishes.
        if (this.isProcessingTasksChanged) {
            this.pendingTasksSnapshot = allTasks;
            return;
        }

        this.isProcessingTasksChanged = true;
        try {
            await this.processTasksChanged(allTasks);

            // Process the latest buffered snapshot if one arrived during our run.
            while (this.pendingTasksSnapshot !== null) {
                const pending = this.pendingTasksSnapshot;
                this.pendingTasksSnapshot = null;
                await this.processTasksChanged(pending);
            }
        } finally {
            this.isProcessingTasksChanged = false;
        }
    }

    /**
     * Core diff-and-push logic. Only ever called from onTasksChanged, which
     * ensures this never runs concurrently with itself.
     */
    private async processTasksChanged(allTasks: Task[]): Promise<void> {
        const settings = this.getSettings();
        const newMap = new Map<string, Task>();

        for (const task of allTasks) {
            if (task.id) {
                newMap.set(task.id, task);
            } else if (settings.autoAssignTaskIds) {
                const key = this.compositeKey(task);
                const restoredId = this.assignedIdAtKey.get(key);
                if (restoredId && !this.pendingIdAssignments.has(key)) {
                    // A previous autoAssignId wrote this ID but the line was
                    // subsequently overwritten without it (e.g. by the edit modal
                    // reading a stale editor buffer).  Re-write with the same ID
                    // so the Planner index mapping stays valid.
                    this.pendingIdAssignments.add(key);
                    const restoredTask = new TaskClass({ ...task, id: restoredId });
                    try {
                        await replaceTaskWithTasks({ originalTask: task, newTasks: restoredTask });
                    } catch (err) {
                        console.error('Planner: restoreId failed', err);
                        this.pendingIdAssignments.delete(key);
                    }
                } else {
                    await this.autoAssignId(task);
                }
            }
        }

        // Deleted tasks: were in previous snapshot and in index, now gone
        for (const [id, _oldTask] of this.previousTasks) {
            if (!newMap.has(id) && this.index.isLinked(id)) {
                await this.pushDelete(id);
            }
        }

        // Created or updated tasks
        for (const [id, newTask] of newMap) {
            const compositeKey = this.compositeKey(newTask);
            this.pendingIdAssignments.delete(compositeKey);
            this.assignedIdAtKey.delete(compositeKey); // ID confirmed in cache — no longer need recovery mapping

            const oldTask = this.previousTasks.get(id);

            if (!oldTask) {
                if (!this.index.isLinked(id)) {
                    // Pre-existing vault task that just received its auto-assigned ID —
                    // do not push to Planner; let the user import explicitly if needed.
                    if (this.preExistingTaskKeys.has(compositeKey)) {
                        this.preExistingTaskKeys.delete(compositeKey);
                        // fall through to rebuildSnapshot so future edits are tracked
                    } else {
                        // Genuinely new task — skip if already closed (done / cancelled)
                        // so that old completed tasks in the vault are never pushed to Planner.
                        const isClosed =
                            newTask.status.type === StatusType.DONE ||
                            newTask.status.type === StatusType.CANCELLED;

                        if (!isClosed) {
                            const pendingPlannerId = this.pendingLinks.get(newTask.descriptionWithoutTags.trim());
                            if (pendingPlannerId) {
                                this.pendingLinks.delete(newTask.descriptionWithoutTags.trim());
                                await this.linkExistingPlannerTask(id, pendingPlannerId);
                            } else {
                                await this.pushCreate(id, newTask);
                            }
                        }
                    }
                }
            } else if (!oldTask.identicalTo(newTask)) {
                if (this.index.isLinked(id)) {
                    await this.pushUpdate(id, oldTask, newTask);
                }
            }
        }

        this.rebuildSnapshot(allTasks);
    }

    private rebuildSnapshot(tasks: Task[]): void {
        this.previousTasks.clear();
        for (const t of tasks) {
            if (t.id) this.previousTasks.set(t.id, t);
        }
    }

    private compositeKey(task: Task): string {
        return `${task.path}:${task.taskLocation.lineNumber}`;
    }

    // -----------------------------------------------------------------------
    // Individual push operations
    // -----------------------------------------------------------------------

    private async pushCreate(tasksId: string, task: Task): Promise<void> {
        const settings = this.getSettings();
        const spec = PlannerTaskMapper.toSpec(task, settings);

        try {
            const { task: pt, etag } = await this.client.createTask(
                spec.planId,
                spec.bucketId,
                spec.title,
                spec.dueDateTime,
                spec.priority,
                spec.percentComplete,
                settings.userId,
            );
            this.index.upsert(tasksId, {
                plannerId: pt.id,
                planId: spec.planId,
                bucketId: spec.bucketId,
                etag,
                lastSyncedAt: Date.now(),
                snapshot: PlannerTaskMapper.toSnapshot(task, settings),
            });
            await this.index.save();
        } catch (err) {
            console.error('Planner: pushCreate failed', err);
            this.queue.enqueue({
                type: 'create',
                tasksId,
                payload: spec as unknown as Record<string, unknown>,
            });
            await this.queue.save();
        }
    }

    private async pushUpdate(tasksId: string, _oldTask: Task, newTask: Task): Promise<void> {
        const entry = this.index.getByTasksId(tasksId);
        if (!entry) return;

        const settings = this.getSettings();

        // Cancellation in Obsidian → delete in Planner
        if (newTask.status.type === StatusType.CANCELLED) {
            await this.pushDelete(tasksId);
            return;
        }

        const patch = PlannerTaskMapper.diffForPatch(newTask, entry.snapshot, settings);
        if (!patch) return;

        try {
            const newEtag = await this.client.updateTask(entry.plannerId, entry.etag, patch);
            this.index.updateEtagAndSnapshot(tasksId, newEtag, PlannerTaskMapper.toSnapshot(newTask, settings));
            await this.index.save();
        } catch (err) {
            console.error('Planner: pushUpdate failed', err);
            this.queue.enqueue({
                type: 'update',
                tasksId,
                plannerId: entry.plannerId,
                etag: entry.etag,
                payload: patch as Record<string, unknown>,
            });
            await this.queue.save();
        }
    }

    private async pushDelete(tasksId: string): Promise<void> {
        const entry = this.index.getByTasksId(tasksId);
        if (!entry) return;

        try {
            await this.client.deleteTask(entry.plannerId, entry.etag);
            this.index.remove(tasksId);
            await this.index.save();
        } catch (err) {
            console.error('Planner: pushDelete failed', err);
            this.queue.enqueue({
                type: 'delete',
                tasksId,
                plannerId: entry.plannerId,
                etag: entry.etag,
            });
            await this.queue.save();
        }
    }

    // -----------------------------------------------------------------------
    // Pull sync  (Planner → Obsidian)
    // -----------------------------------------------------------------------

    /** Pull and apply Planner state for all linked tasks in the given file. */
    async pullSyncFile(file: TFile): Promise<void> {
        if (!this.isReady()) return;

        const fileTasks = this.plugin.getTasks().filter((t) => t.path === file.path && this.index.isLinked(t.id));

        for (const task of fileTasks) {
            await this.pullSyncTask(task);
        }
    }

    private async pullSyncTask(task: Task): Promise<void> {
        const entry = this.index.getByTasksId(task.id);
        if (!entry) return;

        try {
            const { task: pt, etag } = await this.client.getTask(entry.plannerId);
            const { statusSymbol, dueDate } = PlannerTaskMapper.fromPlanner(pt);

            const newStatus = StatusRegistry.getInstance().bySymbolOrCreate(statusSymbol);
            const statusChanged = newStatus.symbol !== task.status.symbol;
            const dueDateChanged =
                (dueDate === null) !== (task.dueDate === null) ||
                (dueDate !== null && task.dueDate !== null && !dueDate.isSame(task.dueDate, 'day'));

            if (statusChanged || dueDateChanged) {
                const updatedTask = new TaskClass({ ...task, status: newStatus, dueDate });
                await replaceTaskWithTasks({ originalTask: task, newTasks: updatedTask });
            }

            const settings = this.getSettings();
            const snapshotTask = statusChanged || dueDateChanged ? new TaskClass({ ...task, status: newStatus, dueDate }) : task;
            this.index.updateEtagAndSnapshot(
                task.id,
                etag,
                PlannerTaskMapper.toSnapshot(snapshotTask, settings),
            );
            await this.index.save();
        } catch (err) {
            const message = err instanceof Error ? err.message : String(err);
            if (message.includes('404')) {
                // Task was deleted in Planner — cancel it in Obsidian and remove from index.
                await this.cancelDeletedTask(task, entry.plannerId);
            } else {
                console.error(`Planner: pullSyncTask failed for plannerId ${entry.plannerId}`, err);
            }
        }
    }

    private async cancelDeletedTask(task: Task, plannerId: string): Promise<void> {
        try {
            const cancelledStatus = StatusRegistry.getInstance().bySymbolOrCreate('-');
            const updatedTask = new TaskClass({ ...task, status: cancelledStatus });
            await replaceTaskWithTasks({ originalTask: task, newTasks: updatedTask });
            this.index.remove(task.id);
            await this.index.save();
            new Notice(`Planner: task cancelled — deleted from Planner: "${task.descriptionWithoutTags.trim()}"`);
        } catch (err) {
            console.error(`Planner: failed to cancel task after Planner 404 (plannerId ${plannerId})`, err);
        }
    }

    // -----------------------------------------------------------------------
    // Offline queue flush
    // -----------------------------------------------------------------------

    /** Replay queued operations. Stops at the first failure to preserve order. */
    async flushQueue(): Promise<void> {
        if (!this.isReady()) return;

        const settings = this.getSettings();

        while (!this.queue.isEmpty()) {
            const op = this.queue.peek();
            if (!op) break;

            try {
                if (op.type === 'create' && op.payload) {
                    const p = op.payload as {
                        planId: string;
                        bucketId: string;
                        title: string;
                        dueDateTime: string | null;
                        priority: number;
                        percentComplete: number;
                    };
                    const { task: pt, etag } = await this.client.createTask(
                        p.planId,
                        p.bucketId,
                        p.title,
                        p.dueDateTime,
                        p.priority,
                        p.percentComplete,
                        settings.userId,
                    );
                    // Rebuild snapshot from the task if still in memory
                    const liveTask = this.previousTasks.get(op.tasksId);
                    this.index.upsert(op.tasksId, {
                        plannerId: pt.id,
                        planId: p.planId,
                        bucketId: p.bucketId,
                        etag,
                        lastSyncedAt: Date.now(),
                        snapshot: liveTask
                            ? PlannerTaskMapper.toSnapshot(liveTask, settings)
                            : { title: p.title, dueDate: p.dueDateTime?.slice(0, 10) ?? null, priority: '3', statusSymbol: ' ', bucketId: p.bucketId },
                    });
                    await this.index.save();
                } else if (op.type === 'update' && op.plannerId && op.etag && op.payload) {
                    const newEtag = await this.client.updateTask(
                        op.plannerId,
                        op.etag,
                        op.payload as Parameters<PlannerApiClient['updateTask']>[2],
                    );
                    const liveTask = this.previousTasks.get(op.tasksId);
                    if (liveTask) {
                        this.index.updateEtagAndSnapshot(
                            op.tasksId,
                            newEtag,
                            PlannerTaskMapper.toSnapshot(liveTask, settings),
                        );
                    }
                    await this.index.save();
                } else if (op.type === 'delete' && op.plannerId && op.etag) {
                    await this.client.deleteTask(op.plannerId, op.etag);
                    this.index.remove(op.tasksId);
                    await this.index.save();
                }

                this.queue.dequeue();
                await this.queue.save();
            } catch (err) {
                console.error('Planner: queue flush failed, will retry later', err);
                break;
            }
        }
    }

    // -----------------------------------------------------------------------
    // Auto ID assignment
    // -----------------------------------------------------------------------

    private async autoAssignId(task: Task): Promise<void> {
        const key = this.compositeKey(task);
        if (this.pendingIdAssignments.has(key)) return;

        this.pendingIdAssignments.add(key);

        const existingIds = Array.from(this.previousTasks.keys());
        const newId = generateUniqueId(existingIds);

        // Record the chosen ID BEFORE the async write.  If the edit modal
        // subsequently overwrites the line without the ID (because it read a stale
        // editor buffer), processTasksChanged will restore this ID instead of
        // generating a second one.
        this.assignedIdAtKey.set(key, newId);

        const updatedTask = new TaskClass({ ...task, id: newId });

        try {
            await replaceTaskWithTasks({ originalTask: task, newTasks: updatedTask });
        } catch (err) {
            console.error('Planner: autoAssignId failed', err);
            this.pendingIdAssignments.delete(key);
            this.assignedIdAtKey.delete(key);
        }
        // pendingIdAssignments / assignedIdAtKey entries are removed in
        // processTasksChanged when the confirmed-ID task appears in newMap.
    }

    // -----------------------------------------------------------------------
    // UC2: link an existing Planner task to a newly created Obsidian task
    // -----------------------------------------------------------------------

    /**
     * Called from the Create/Edit modal when the user selects an existing
     * Planner task to link instead of creating a new one.
     */
    async linkExistingPlannerTask(tasksId: string, plannerId: string): Promise<void> {
        const settings = this.getSettings();
        const { task: pt, etag } = await this.client.getTask(plannerId);
        const liveTask = this.previousTasks.get(tasksId);
        this.index.upsert(tasksId, {
            plannerId,
            planId: pt.planId,
            bucketId: pt.bucketId,
            etag,
            lastSyncedAt: Date.now(),
            snapshot: liveTask
                ? PlannerTaskMapper.toSnapshot(liveTask, settings)
                : { title: pt.title, dueDate: pt.dueDateTime?.slice(0, 10) ?? null, priority: '3', statusSymbol: ' ', bucketId: pt.bucketId },
        });
        await this.index.save();
    }

    /**
     * UC2: store a pending Planner link by description so onTasksChanged can
     * match the next new task with that description and link it instead of
     * creating a duplicate.
     */
    storePendingLink(plannerId: string, description: string): void {
        this.pendingLinks.set(description, plannerId);
    }

    // -----------------------------------------------------------------------
    // UC8: Bulk import from Planner
    // -----------------------------------------------------------------------

    /**
     * Import active Planner tasks into an Obsidian markdown file.
     *
     * @param planId       Plan to import from.
     * @param bucketId     Limit to a specific bucket, or '' for all buckets.
     * @param targetPath   Vault-relative path of the file to append tasks to
     *                     (created if it doesn't exist).
     * @returns            Number of tasks imported.
     */
    async importFromPlanner(planId: string, bucketId: string | undefined, targetPath: string): Promise<number> {
        const allTasks = await this.client.getPlanTasks(planId);

        // If a specific bucket was requested use it; otherwise fall back to all watched buckets.
        const watchedIds = allWatchedBucketIds(this.getSettings());
        const tasks = allTasks.filter((t) => {
            if (t.percentComplete >= 100) return false;
            if (bucketId) return t.bucketId === bucketId;
            if (watchedIds.length > 0) return watchedIds.includes(t.bucketId);
            return true; // no filter at all
        });

        if (tasks.length === 0) return 0;

        const vault = this.plugin.app.vault;
        const normalizedPath = normalizePath(targetPath);

        // Ensure file exists
        let file = vault.getFileByPath(normalizedPath);
        if (!file) {
            file = await vault.create(normalizedPath, '');
        }

        const lines: string[] = [];
        const existingIds = Array.from(this.previousTasks.keys());
        const usedIds = new Set(existingIds);

        // Build a bucket-id → tag lookup from the configured tag mappings
        const settings = this.getSettings();
        const bucketTagMap = new Map<string, string>();
        for (const m of settings.tagMappings) {
            if (m.bucketId && m.tag) bucketTagMap.set(m.bucketId, m.tag);
        }

        for (const pt of tasks) {
            // Skip tasks already in the index
            if (this.index.getByPlannerId(pt.id)) continue;

            const statusSymbol = pt.percentComplete >= 100 ? 'x' : pt.percentComplete > 0 ? '/' : ' ';
            const dueStr = pt.dueDateTime ? ` 📅 ${pt.dueDateTime.slice(0, 10)}` : '';
            const newId = generateUniqueId([...usedIds]);
            usedIds.add(newId);

            // Append the tag that maps to this bucket, if any
            const tag = bucketTagMap.get(pt.bucketId);
            const tagStr = tag ? ` ${tag}` : '';

            const line = `- [${statusSymbol}] ${pt.title}${tagStr}${dueStr} 🆔 ${newId}`;
            lines.push(line);

            // Pre-register in index so push sync doesn't try to create it again
            this.index.upsert(newId, {
                plannerId: pt.id,
                planId: pt.planId,
                bucketId: pt.bucketId,
                etag: '', // will be refreshed on first pull sync
                lastSyncedAt: Date.now(),
                snapshot: {
                    title: pt.title,
                    dueDate: pt.dueDateTime?.slice(0, 10) ?? null,
                    priority: String(pt.priority),
                    statusSymbol,
                    bucketId: pt.bucketId,
                },
            });
        }

        if (lines.length === 0) return 0;

        const existing = await vault.read(file);
        const separator = existing.length > 0 && !existing.endsWith('\n') ? '\n' : '';
        const header = `\n## Imported from Planner — ${new Date().toLocaleDateString()}\n\n`;
        await vault.modify(file, existing + separator + header + lines.join('\n') + '\n');
        await this.index.save();

        new Notice(`Planner: imported ${lines.length} task(s) into ${normalizedPath}`);
        return lines.length;
    }

    /** Expose the API client for settings-UI calls (loading plan/bucket lists). */
    getApiClient(): PlannerApiClient {
        return this.client;
    }

    /** Show a non-dismissable notice on sync errors (used by settings UI). */
    showError(message: string): void {
        new Notice(`Planner sync: ${message}`, 8000);
    }
}
