import { debounce } from 'obsidian';
import { TaskModal, type TaskModalParams } from '../Obsidian/TaskModal';
import type { Task } from '../Task/Task';
import type { PlannerSyncManager } from './PlannerSyncManager';
import type { PlannerTask } from './PlannerApiClient';
import { getSettings } from '../Config/Settings';
import { allWatchedBucketIds } from './PlannerSettings';

// ---------------------------------------------------------------------------
// PlannerTaskModal
// ---------------------------------------------------------------------------

/**
 * Extends the standard Create/Edit modal so that typing in the Description
 * field searches pre-fetched Planner tasks inline.
 *
 * A suggestion panel is inserted directly below the description textarea.
 * Selecting a result links the Obsidian task to an existing Planner task
 * instead of creating a duplicate.  Leaving the field without a selection
 * lets the normal sync machinery create a new Planner task.
 */
export class PlannerTaskModal extends TaskModal {
    private readonly syncManager: PlannerSyncManager;

    private selectedPlannerId: string | null = null;
    private selectedPlannerTitle: string | null = null;

    private plannerTasks: PlannerTask[] = [];
    private tasksLoaded = false;

    // DOM refs populated by buildSuggestPanel()
    private suggestStatusEl: HTMLElement | null = null;
    private suggestListEl: HTMLElement | null = null;
    private suggestBadgeEl: HTMLElement | null = null;

    constructor(params: TaskModalParams & { syncManager: PlannerSyncManager }) {
        const holder: { modal: PlannerTaskModal | null } = { modal: null };

        super({
            app: params.app,
            task: params.task,
            allTasks: params.allTasks,
            onSubmit: (updatedTasks: Task[]) => {
                params.onSubmit(updatedTasks);
                holder.modal?.handlePlannerOnSubmit(updatedTasks);
            },
        });

        this.syncManager = params.syncManager;
        holder.modal = this;
    }

    // -----------------------------------------------------------------------
    // onOpen
    // -----------------------------------------------------------------------

    public onOpen(): void {
        super.onOpen();

        const ps = getSettings().plannerSettings;
        if (!ps.enabled || !ps.defaultPlanId) return;

        this.buildSuggestPanel();
        this.fetchPlannerTasks();
    }

    // -----------------------------------------------------------------------
    // Suggest panel — inserted directly below the description textarea
    // -----------------------------------------------------------------------

    private buildSuggestPanel(): void {
        const { contentEl } = this;

        const descSection = contentEl.querySelector('.tasks-modal-description-section') as HTMLElement | null;
        const textarea = contentEl.querySelector<HTMLTextAreaElement>('textarea#description');
        if (!descSection || !textarea) return;

        // Wrapper inserted right after the description section
        const panel = document.createElement('div');
        panel.className = 'planner-suggest-panel';
        panel.style.cssText =
            'padding:2px 0 6px; border-bottom:1px solid var(--background-modifier-border);';
        descSection.insertAdjacentElement('afterend', panel);

        // Status / hint line
        const statusEl = panel.createEl('p', {});
        statusEl.style.cssText =
            'font-size:0.78em; color:var(--text-muted); margin:2px 0 4px; padding: 0 2px;';
        statusEl.textContent = 'Loading Planner tasks…';

        // Suggestion list
        const listEl = panel.createEl('ul', {});
        listEl.style.cssText =
            'list-style:none; margin:0; padding:0; max-height:150px; overflow-y:auto;';

        // Selected-task badge (shown after the user picks a suggestion)
        const badgeEl = panel.createDiv({});
        badgeEl.style.cssText =
            'display:none; font-size:0.82em; padding:3px 8px; margin-top:4px; border-radius:4px;' +
            ' background:var(--background-modifier-success-hover,rgba(0,200,100,0.15));' +
            ' color:var(--text-success, var(--text-normal)); display:flex; gap:6px; align-items:center;';

        const badgeText = badgeEl.createSpan();
        const clearBtn = badgeEl.createEl('button', { text: '✕' });
        clearBtn.style.cssText =
            'background:none; border:none; cursor:pointer; padding:0 2px; font-size:0.9em; color:inherit;';
        clearBtn.title = 'Unlink — will create a new Planner task on submit';
        clearBtn.addEventListener('click', () => {
            this.clearSelection(statusEl, listEl, badgeEl);
            textarea.focus();
        });
        badgeEl.style.display = 'none'; // hide initially

        this.suggestStatusEl = statusEl;
        this.suggestListEl = listEl;
        this.suggestBadgeEl = badgeEl;

        // Listen to description field: search as the user types
        const doSearch = debounce(() => {
            if (this.selectedPlannerId) {
                // User edited the field after linking — unlink silently
                this.clearSelection(statusEl, listEl, badgeEl);
            }
            this.renderSuggestions(textarea.value, listEl, statusEl, textarea, badgeEl, badgeText);
        }, 200, true);

        textarea.addEventListener('input', doSearch);
    }

    // -----------------------------------------------------------------------
    // Suggestion rendering
    // -----------------------------------------------------------------------

    private renderSuggestions(
        query: string,
        listEl: HTMLElement,
        statusEl: HTMLElement,
        textarea: HTMLTextAreaElement,
        badgeEl: HTMLElement,
        badgeText: HTMLElement,
    ): void {
        listEl.empty();

        if (!query.trim()) {
            statusEl.textContent = this.tasksLoaded
                ? `${this.plannerTasks.length} Planner tasks loaded — type to search`
                : 'Loading Planner tasks…';
            return;
        }

        if (!this.tasksLoaded) {
            statusEl.textContent = 'Still loading Planner tasks…';
            return;
        }

        const lower = query.toLowerCase();
        const matches = this.plannerTasks
            .filter((t) => t.title.toLowerCase().includes(lower))
            .slice(0, 8);

        if (matches.length === 0) {
            statusEl.textContent = 'No match — a new Planner task will be created on save.';
            return;
        }

        statusEl.textContent = '';

        for (const pt of matches) {
            const li = listEl.createEl('li');
            li.style.cssText =
                'padding:4px 8px; cursor:pointer; border-radius:4px; font-size:0.88em;' +
                ' white-space:nowrap; overflow:hidden; text-overflow:ellipsis;';
            li.title = pt.title;

            const icon = pt.percentComplete >= 100 ? '✅' : pt.percentComplete > 0 ? '🔄' : '⬜';
            li.textContent = `${icon} ${pt.title}`;

            li.addEventListener('mouseenter', () => (li.style.background = 'var(--background-modifier-hover)'));
            li.addEventListener('mouseleave', () => (li.style.background = ''));
            li.addEventListener('click', () => {
                this.selectedPlannerId = pt.id;
                this.selectedPlannerTitle = pt.title;

                listEl.empty();
                statusEl.textContent = '';

                badgeText.textContent = `🔗 Linked to: "${pt.title}"`;
                badgeEl.style.display = 'flex';

                textarea.disabled = true;
            });

            listEl.appendChild(li);
        }
    }

    private clearSelection(statusEl: HTMLElement, listEl: HTMLElement, badgeEl: HTMLElement): void {
        this.selectedPlannerId = null;
        this.selectedPlannerTitle = null;
        badgeEl.style.display = 'none';
        listEl.empty();

        const textarea = this.contentEl.querySelector<HTMLTextAreaElement>('textarea#description');
        if (textarea) textarea.disabled = false;

        statusEl.textContent = this.tasksLoaded
            ? `${this.plannerTasks.length} Planner tasks loaded — type to search`
            : 'Loading Planner tasks…';
    }

    // -----------------------------------------------------------------------
    // Background task fetch
    // -----------------------------------------------------------------------

    private fetchPlannerTasks(): void {
        const client = this.syncManager.getApiClient();
        const ps = getSettings().plannerSettings;

        const planIds = new Set<string>([ps.defaultPlanId]);
        for (const m of ps.tagMappings) {
            if (m.planId) planIds.add(m.planId);
        }

        const watchedIds = new Set(allWatchedBucketIds(ps));

        const fetches = Array.from(planIds).map((id) =>
            client.getPlanTasks(id).catch(() => [] as PlannerTask[]),
        );

        Promise.all(fetches).then((results) => {
            const all = results.flat().filter((t) => t.percentComplete < 100);
            this.plannerTasks =
                watchedIds.size > 0 ? all.filter((t) => watchedIds.has(t.bucketId)) : all;
            this.tasksLoaded = true;

            if (this.suggestStatusEl && !this.selectedPlannerId) {
                const textarea = this.contentEl.querySelector<HTMLTextAreaElement>('textarea#description');
                const hasQuery = textarea && textarea.value.trim().length > 0;
                if (!hasQuery) {
                    this.suggestStatusEl.textContent =
                        `${this.plannerTasks.length} Planner tasks loaded — type to search`;
                }
            }
        });
    }

    // -----------------------------------------------------------------------
    // Submit handler
    // -----------------------------------------------------------------------

    private handlePlannerOnSubmit(updatedTasks: Task[]): void {
        if (updatedTasks.length === 0) return;
        const task = updatedTasks[0];

        if (this.selectedPlannerId) {
            if (task.id) {
                this.syncManager.linkExistingPlannerTask(task.id, this.selectedPlannerId).catch(console.error);
            } else {
                this.syncManager.storePendingLink(this.selectedPlannerId, task.descriptionWithoutTags.trim());
            }
        }
        // No selection → normal push-create via onTasksChanged
    }
}
