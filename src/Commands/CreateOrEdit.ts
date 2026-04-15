import { App, Editor, MarkdownView, View } from 'obsidian';
import { TaskModal } from '../Obsidian/TaskModal';
import type { Task } from '../Task/Task';
import { DateFallback } from '../DateTime/DateFallback';
import { taskFromLine } from './CreateOrEditTaskParser';
import { getSettings } from '../Config/Settings';
import { PlannerAuth } from '../Planner/PlannerAuth';
import { PlannerTaskModal } from '../Planner/PlannerTaskModal';
import type { PlannerSyncManager } from '../Planner/PlannerSyncManager';

export const createOrEdit = (
    checking: boolean,
    editor: Editor,
    view: View,
    app: App,
    allTasks: Task[],
    plannerSyncManager?: PlannerSyncManager,
) => {
    if (checking) {
        return view instanceof MarkdownView;
    }

    if (!(view instanceof MarkdownView)) {
        // Should never happen due to check above.
        return;
    }

    const path = view.file?.path;
    if (path === undefined) {
        return;
    }

    const cursorPosition = editor.getCursor();
    const lineNumber = cursorPosition.line;
    const line = editor.getLine(lineNumber);
    const task = taskFromLine({ line, path });

    const onSubmit = (updatedTasks: Task[]): void => {
        const serialized = DateFallback.removeInferredStatusIfNeeded(task, updatedTasks)
            .map((task: Task) => task.toFileLineString())
            .join('\n');
        editor.setLine(lineNumber, serialized);
    };

    // Use PlannerTaskModal when the Planner integration is active.
    const ps = getSettings().plannerSettings;
    const usePlanner = plannerSyncManager && ps.enabled && PlannerAuth.isAuthenticated(ps) && ps.defaultPlanId !== '';

    const taskModal = usePlanner
        ? new PlannerTaskModal({ app, task, onSubmit, allTasks, syncManager: plannerSyncManager! })
        : new TaskModal({ app, task, onSubmit, allTasks });

    taskModal.open();
};
