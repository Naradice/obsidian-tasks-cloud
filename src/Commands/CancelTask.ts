import { Editor, type MarkdownFileInfo, MarkdownView } from 'obsidian';
import { TasksFile } from '../Scripting/TasksFile';
import { StatusRegistry } from '../Statuses/StatusRegistry';
import { Task } from '../Task/Task';
import { TaskLocation } from '../Task/TaskLocation';
import { TaskRegularExpressions } from '../Task/TaskRegularExpressions';

/**
 * Editor command: cancel the task on the cursor line.
 *
 * - If the line is a recognised Task, replaces its status with the first
 *   status whose symbol is '-' (cancelled).
 * - If the line is a plain checklist item (not parsed as a Task), sets
 *   its checkbox to `[-]`.
 * - Has no effect on non-list lines (command is still enabled so the
 *   user can assign a hotkey unconditionally).
 */
export const cancelTask = (checking: boolean, editor: Editor, view: MarkdownView | MarkdownFileInfo): boolean => {
    if (checking) {
        return view instanceof MarkdownView;
    }

    if (!(view instanceof MarkdownView)) {
        return false;
    }

    const path = view.file?.path;
    if (path === undefined) return false;

    const cursorPos = editor.getCursor();
    const lineNumber = cursorPos.line;
    const line = editor.getLine(lineNumber);

    const task = Task.fromLine({
        line,
        taskLocation: TaskLocation.fromUnknownPosition(new TasksFile(path)),
        fallbackDate: null,
    });

    if (task !== null) {
        const cancelledStatus = StatusRegistry.getInstance().bySymbolOrCreate('-');
        const cancelledTask = new Task({ ...task, status: cancelledStatus });
        editor.setLine(lineNumber, cancelledTask.toFileLineString());
        return true;
    }

    // Fall back: toggle a bare checklist item to [-]
    if (TaskRegularExpressions.taskRegex.test(line)) {
        editor.setLine(lineNumber, line.replace(TaskRegularExpressions.taskRegex, '$1$2 [-] $4'));
    }

    return true;
};
