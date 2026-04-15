import type { Task } from '../Task/Task';
import { Priority } from '../Task/Priority';
import { StatusType } from '../Statuses/StatusConfiguration';
import type { PlannerTask } from './PlannerApiClient';
import type { PlannerPriority, PlannerSettings } from './PlannerSettings';
import type { SyncSnapshot } from './PlannerTaskIndex';

// ---------------------------------------------------------------------------
// Intermediate spec used when creating / comparing tasks
// ---------------------------------------------------------------------------

export interface PlannerTaskSpec {
    planId: string;
    bucketId: string;
    title: string;
    dueDateTime: string | null; // ISO 8601 UTC
    priority: PlannerPriority;
    percentComplete: number;
}

// ---------------------------------------------------------------------------
// Mapper
// ---------------------------------------------------------------------------

export class PlannerTaskMapper {
    /**
     * Resolve which plan+bucket a task belongs to.
     * Iterates through the task's tags and returns the first matching mapping.
     * Falls back to the configured default plan/bucket.
     */
    static resolvePlanBucket(task: Task, settings: PlannerSettings): { planId: string; bucketId: string } {
        for (const tag of task.tags) {
            const mapping = settings.tagMappings.find((m) => m.tag === tag);
            if (mapping) return { planId: mapping.planId, bucketId: mapping.bucketId };
        }
        return { planId: settings.defaultPlanId, bucketId: settings.defaultBucketId };
    }

    /** Build the full Planner spec from an Obsidian task. */
    static toSpec(task: Task, settings: PlannerSettings): PlannerTaskSpec {
        const { planId, bucketId } = PlannerTaskMapper.resolvePlanBucket(task, settings);
        return {
            planId,
            bucketId,
            title: task.descriptionWithoutTags.trim(),
            dueDateTime: task.dueDate?.isValid() ? task.dueDate.utc().startOf('day').toISOString() : null,
            priority: PlannerTaskMapper.mapPriority(task.priority, settings),
            percentComplete: PlannerTaskMapper.statusToPercent(task),
        };
    }

    /** Build a snapshot for the index from the current task state. */
    static toSnapshot(task: Task, settings: PlannerSettings): SyncSnapshot {
        const { bucketId } = PlannerTaskMapper.resolvePlanBucket(task, settings);
        return {
            title: task.descriptionWithoutTags.trim(),
            dueDate: task.dueDate?.isValid() ? task.dueDate.format('YYYY-MM-DD') : null,
            priority: task.priority,
            statusSymbol: task.status.symbol,
            bucketId,
        };
    }

    /**
     * Compare the current task against the stored snapshot to produce only the
     * changed fields. Returns null when nothing Planner-relevant has changed.
     */
    static diffForPatch(
        task: Task,
        snapshot: SyncSnapshot,
        settings: PlannerSettings,
    ): Partial<Pick<PlannerTask, 'title' | 'dueDateTime' | 'priority' | 'percentComplete' | 'bucketId'>> | null {
        const spec = PlannerTaskMapper.toSpec(task, settings);
        const patch: Record<string, unknown> = {};

        if (spec.title !== snapshot.title) patch.title = spec.title;

        const newDue = task.dueDate?.isValid() ? task.dueDate.format('YYYY-MM-DD') : null;
        if (newDue !== snapshot.dueDate) patch.dueDateTime = spec.dueDateTime;

        if (task.priority !== snapshot.priority) patch.priority = spec.priority;

        const newPercent = PlannerTaskMapper.statusToPercent(task);
        const oldPercent = PlannerTaskMapper.symbolToPercent(snapshot.statusSymbol);
        if (newPercent !== oldPercent) patch.percentComplete = newPercent;

        if (spec.bucketId !== snapshot.bucketId) patch.bucketId = spec.bucketId;

        return Object.keys(patch).length > 0
            ? (patch as Partial<Pick<PlannerTask, 'title' | 'dueDateTime' | 'priority' | 'percentComplete' | 'bucketId'>>)
            : null;
    }

    /**
     * Map a Planner task back to changes for an Obsidian Task.
     * Returns the new status symbol and due date so the caller can create an
     * updated Task with  new Task({ ...existing, status, dueDate }).
     */
    static fromPlanner(plannerTask: PlannerTask): { statusSymbol: string; dueDate: moment.Moment | null } {
        return {
            statusSymbol: PlannerTaskMapper.percentToSymbol(plannerTask.percentComplete),
            dueDate: plannerTask.dueDateTime ? window.moment(plannerTask.dueDateTime).local() : null,
        };
    }

    // -----------------------------------------------------------------------
    // Private helpers
    // -----------------------------------------------------------------------

    static statusToPercent(task: Task): number {
        switch (task.status.type) {
            case StatusType.DONE:
                return 100;
            case StatusType.IN_PROGRESS:
                return 50;
            default:
                return 0;
        }
    }

    private static symbolToPercent(symbol: string): number {
        if (symbol === 'x' || symbol === 'X') return 100;
        if (symbol === '/') return 50;
        return 0;
    }

    private static percentToSymbol(percent: number): string {
        if (percent >= 100) return 'x';
        if (percent > 0) return '/';
        return ' ';
    }

    static mapPriority(priority: Priority, settings: PlannerSettings): PlannerPriority {
        switch (priority) {
            case Priority.Highest:
                return settings.priorityMapping.highest;
            case Priority.High:
                return settings.priorityMapping.high;
            case Priority.Medium:
                return settings.priorityMapping.medium;
            case Priority.Low:
                return settings.priorityMapping.low;
            case Priority.Lowest:
                return settings.priorityMapping.lowest;
            default:
                return settings.priorityMapping.none;
        }
    }
}
