import { Notice, PluginSettingTab, Setting, debounce } from 'obsidian';
import { StatusConfiguration, StatusType } from '../Statuses/StatusConfiguration';
import type TasksPlugin from '../main';
import { StatusRegistry } from '../Statuses/StatusRegistry';
import { Status } from '../Statuses/Status';
import type { StatusCollection } from '../Statuses/StatusCollection';
import { createStatusRegistryReport } from '../Statuses/StatusRegistryReport';
import { i18n } from '../i18n/i18n';
import type { TasksEvents } from '../Obsidian/TasksEvents';
import * as Themes from './Themes';
import {
    type HeadingState,
    TASK_FORMATS,
    getSettings,
    isFeatureEnabled,
    updateGeneralSetting,
    updateSettings,
} from './Settings';
import { GlobalFilter } from './GlobalFilter';
import { StatusSettings } from './StatusSettings';

import { CustomStatusModal } from './CustomStatusModal';
import { GlobalQuery } from './GlobalQuery';
import { PresetsSettingsUI } from './PresetsSettingsUI';
import { PlannerAuth } from '../Planner/PlannerAuth';
import type { PlannerSyncManager } from '../Planner/PlannerSyncManager';
import type { PlannerBucket, PlannerPlan } from '../Planner/PlannerApiClient';
import type { PlannerPriority, TagBucketMapping } from '../Planner/PlannerSettings';
import { Priority } from '../Task/Priority';

export class SettingsTab extends PluginSettingTab {
    // If the UI needs a more complex setting you can create a
    // custom function and specify it from the json file. It will
    // then be rendered instead of a normal checkbox or text box.
    customFunctions: { [K: string]: Function } = {
        insertTaskCoreStatusSettings: this.insertTaskCoreStatusSettings.bind(this),
        insertCustomTaskStatusSettings: this.insertCustomTaskStatusSettings.bind(this),
    };

    private readonly plugin: TasksPlugin;
    private readonly presetsSettingsUI;
    private readonly events: TasksEvents;
    private plannerSyncManager: PlannerSyncManager | null = null;

    /** AbortController for an in-progress device code authentication poll. */
    private authAbortController: AbortController | null = null;

    constructor({ plugin, events }: { plugin: TasksPlugin; events: TasksEvents }) {
        super(plugin.app, plugin);

        this.plugin = plugin;
        this.presetsSettingsUI = new PresetsSettingsUI(plugin, events);
        this.events = events;
    }

    /** Called from main.ts once PlannerSyncManager is initialised. */
    setPlannerSyncManager(manager: PlannerSyncManager): void {
        this.plannerSyncManager = manager;
    }

    private static createFragmentWithHTML = (html: string) =>
        createFragment((documentFragment) => (documentFragment.createDiv().innerHTML = html));

    public async saveSettings(update?: boolean): Promise<void> {
        await this.plugin.saveSettings();

        if (update) {
            this.display();
        }
    }

    public display(): void {
        const { containerEl } = this;

        containerEl.empty();
        this.containerEl.addClass('tasks-settings');

        new Setting(containerEl)
            .setName(i18n.t('settings.format.name'))
            .setDesc(
                SettingsTab.createFragmentWithHTML(
                    `<p>${i18n.t('settings.format.description.line1')}</p>` +
                        `<p>${i18n.t('settings.format.description.line2')}</p>` +
                        `<p>${i18n.t('settings.changeRequiresRestart')}</p>` +
                        this.seeTheDocumentation(
                            'https://publish.obsidian.md/tasks/Reference/Task+Formats/About+Task+Formats',
                        ),
                ),
            )
            .addDropdown((dropdown) => {
                for (const key of Object.keys(TASK_FORMATS) as (keyof TASK_FORMATS)[]) {
                    dropdown.addOption(key, TASK_FORMATS[key].getDisplayName());
                }

                dropdown.setValue(getSettings().taskFormat).onChange(async (value) => {
                    updateSettings({ taskFormat: value as keyof TASK_FORMATS });
                    await this.plugin.saveSettings();
                });
            });

        // ---------------------------------------------------------------------------
        new Setting(containerEl).setName(i18n.t('settings.globalFilter.heading')).setHeading();
        // ---------------------------------------------------------------------------
        let globalFilterHidden: Setting | null = null;

        new Setting(containerEl)
            .setName(i18n.t('settings.globalFilter.filter.name'))
            .setDesc(
                SettingsTab.createFragmentWithHTML(
                    `<p><b>${i18n.t('settings.globalFilter.filter.description.line1')}</b></p>` +
                        `<p>${i18n.t('settings.globalFilter.filter.description.line2')}<p>` +
                        `<p>${i18n.t('settings.globalFilter.filter.description.line3')}</br>` +
                        `${i18n.t('settings.globalFilter.filter.description.line4')}</p>` +
                        this.seeTheDocumentation('https://publish.obsidian.md/tasks/Getting+Started/Global+Filter'),
                ),
            )
            .addText((text) => {
                // I wanted to make this say 'for example, #task or TODO'
                // but wasn't able to figure out how to make the text box
                // wide enough for the whole string to be visible.
                text.setPlaceholder(i18n.t('settings.globalFilter.filter.placeholder'))
                    .setValue(GlobalFilter.getInstance().get())
                    .onChange(
                        debounce(
                            async (value) => {
                                updateSettings({ globalFilter: value });
                                GlobalFilter.getInstance().set(value);
                                await this.plugin.saveSettings();
                                setSettingVisibility(globalFilterHidden, value.length > 0);

                                this.events.triggerReloadVault();
                            },
                            500,
                            true,
                        ),
                    );
            });

        globalFilterHidden = new Setting(containerEl)
            .setName(i18n.t('settings.globalFilter.removeFilter.name'))
            .setDesc(
                SettingsTab.createFragmentWithHTML(
                    `<p>${i18n.t('settings.globalFilter.removeFilter.description')}</p>` +
                        `<p>${i18n.t('settings.changeRequiresRestart')}</p>`,
                ),
            )
            .addToggle((toggle) => {
                const settings = getSettings();

                toggle.setValue(settings.removeGlobalFilter).onChange(async (value) => {
                    updateSettings({ removeGlobalFilter: value });
                    GlobalFilter.getInstance().setRemoveGlobalFilter(value);
                    await this.plugin.saveSettings();
                });
            });
        setSettingVisibility(globalFilterHidden, getSettings().globalFilter.length > 0);

        // ---------------------------------------------------------------------------
        new Setting(containerEl).setName(i18n.t('settings.globalQuery.heading')).setHeading();
        // ---------------------------------------------------------------------------

        makeMultilineTextSetting(
            new Setting(containerEl)
                .setDesc(
                    SettingsTab.createFragmentWithHTML(
                        `<p>${i18n.t('settings.globalQuery.query.description')}</p>` +
                            this.seeTheDocumentation('https://publish.obsidian.md/tasks/Queries/Global+Query'),
                    ),
                )
                .addTextArea((text) => {
                    const settings = getSettings();

                    text.inputEl.rows = 4;
                    text.setPlaceholder('# ' + i18n.t('settings.globalQuery.query.placeholder'))
                        .setValue(settings.globalQuery)
                        .onChange(async (value) => {
                            updateSettings({ globalQuery: value });
                            GlobalQuery.getInstance().set(value);
                            await this.plugin.saveSettings();

                            this.events.triggerReloadOpenSearchResults();
                        });
                }),
        );

        // ---------------------------------------------------------------------------
        new Setting(containerEl)
            .setName(i18n.t('settings.presets.name'))
            .setHeading()
            .setDesc(
                SettingsTab.createFragmentWithHTML(
                    '<p>' +
                        i18n.t('settings.presets.line1', {
                            name: '<code>name</code>',
                            instruction1: '<code>preset name</code>',
                            instruction2: '<code>{{preset.name}}</code>',
                        }) +
                        '</p><p>' +
                        i18n.t('settings.presets.line2') +
                        '</p>' +
                        this.seeTheDocumentation('https://publish.obsidian.md/tasks/Queries/Presets'),
                ),
            );
        // ---------------------------------------------------------------------------
        this.presetsSettingsUI.renderPresetsSettings(containerEl);

        // ---------------------------------------------------------------------------
        new Setting(containerEl).setName(i18n.t('settings.statuses.heading')).setHeading();
        // ---------------------------------------------------------------------------

        const { headingOpened } = getSettings();

        // Directly define the JSON data as a constant object
        const settingsJson = [
            {
                text: i18n.t('settings.statuses.coreStatuses.heading'),
                level: 'h3',
                class: '',
                open: true,
                notice: {
                    class: 'setting-item-description',
                    text: null,
                    html:
                        '<p>' +
                        i18n.t('settings.statuses.coreStatuses.description.line1') +
                        '</p><p>' +
                        i18n.t('settings.statuses.coreStatuses.description.line2') +
                        '</p><p>' +
                        i18n.t('settings.changeRequiresRestart') +
                        '</p>',
                },
                settings: [
                    {
                        name: '',
                        description: '',
                        type: 'function',
                        initialValue: '',
                        placeholder: '',
                        settingName: 'insertTaskCoreStatusSettings',
                        featureFlag: '',
                        notice: null,
                    },
                ],
            },
            {
                text: i18n.t('settings.statuses.customStatuses.heading'),
                level: 'h3',
                class: '',
                open: true,
                notice: {
                    class: 'setting-item-description',
                    text: null,
                    html:
                        '<p>' +
                        i18n.t('settings.statuses.customStatuses.description.line1') +
                        '</p><p>' +
                        i18n.t('settings.statuses.customStatuses.description.line2') +
                        '</p><p>' +
                        i18n.t('settings.statuses.customStatuses.description.line3') +
                        '</p><p>' +
                        i18n.t('settings.changeRequiresRestart') +
                        '</p><p></p><p>' +
                        `<a href="https://publish.obsidian.md/tasks/Getting+Started/Statuses">${i18n.t(
                            'settings.statuses.customStatuses.description.line4',
                        )}</a></p>`,
                },
                settings: [
                    {
                        name: '',
                        description: '',
                        type: 'function',
                        initialValue: '',
                        placeholder: '',
                        settingName: 'insertCustomTaskStatusSettings',
                        featureFlag: '',
                        notice: null,
                    },
                ],
            },
        ];

        // Original usage remains unchanged
        settingsJson.forEach((heading) => {
            const initiallyOpen = headingOpened[heading.text] ?? true;
            const detailsContainer = this.addOneSettingsBlock(containerEl, heading, headingOpened);
            detailsContainer.open = initiallyOpen;
        });

        // ---------------------------------------------------------------------------
        new Setting(containerEl).setName(i18n.t('settings.dates.heading')).setHeading();
        // ---------------------------------------------------------------------------

        new Setting(containerEl)
            .setName(i18n.t('settings.dates.createdDate.name'))
            .setDesc(
                SettingsTab.createFragmentWithHTML(
                    i18n.t('settings.dates.createdDate.description') +
                        '</br>' +
                        this.seeTheDocumentation(
                            'https://publish.obsidian.md/tasks/Getting+Started/Dates#Created+date',
                        ),
                ),
            )
            .addToggle((toggle) => {
                const settings = getSettings();
                toggle.setValue(settings.setCreatedDate).onChange(async (value) => {
                    updateSettings({ setCreatedDate: value });
                    await this.plugin.saveSettings();
                });
            });

        new Setting(containerEl)
            .setName(i18n.t('settings.dates.doneDate.name'))
            .setDesc(
                SettingsTab.createFragmentWithHTML(
                    i18n.t('settings.dates.doneDate.description') +
                        '</br>' +
                        this.seeTheDocumentation('https://publish.obsidian.md/tasks/Getting+Started/Dates#Done+date'),
                ),
            )
            .addToggle((toggle) => {
                const settings = getSettings();
                toggle.setValue(settings.setDoneDate).onChange(async (value) => {
                    updateSettings({ setDoneDate: value });
                    await this.plugin.saveSettings();
                });
            });

        new Setting(containerEl)
            .setName(i18n.t('settings.dates.cancelledDate.name'))
            .setDesc(
                SettingsTab.createFragmentWithHTML(
                    i18n.t('settings.dates.cancelledDate.description') +
                        '</br>' +
                        this.seeTheDocumentation(
                            'https://publish.obsidian.md/tasks/Getting+Started/Dates#Cancelled+date',
                        ),
                ),
            )
            .addToggle((toggle) => {
                const settings = getSettings();
                toggle.setValue(settings.setCancelledDate).onChange(async (value) => {
                    updateSettings({ setCancelledDate: value });
                    await this.plugin.saveSettings();
                });
            });

        // ---------------------------------------------------------------------------
        new Setting(containerEl).setName(i18n.t('settings.datesFromFileNames.heading')).setHeading();
        // ---------------------------------------------------------------------------
        let scheduledDateExtraFormat: Setting | null = null;
        let scheduledDateFolders: Setting | null = null;

        new Setting(containerEl)
            .setName(i18n.t('settings.datesFromFileNames.scheduledDate.toggle.name'))
            .setDesc(
                SettingsTab.createFragmentWithHTML(
                    i18n.t('settings.datesFromFileNames.scheduledDate.toggle.description.line1') +
                        '</br>' +
                        i18n.t('settings.datesFromFileNames.scheduledDate.toggle.description.line2') +
                        '</br>' +
                        i18n.t('settings.datesFromFileNames.scheduledDate.toggle.description.line3') +
                        '</br>' +
                        i18n.t('settings.datesFromFileNames.scheduledDate.toggle.description.line4') +
                        '</br>' +
                        `<p>${i18n.t('settings.changeRequiresRestart')}</p>` +
                        this.seeTheDocumentation(
                            'https://publish.obsidian.md/tasks/Getting+Started/Use+Filename+as+Default+Date',
                        ),
                ),
            )
            .addToggle((toggle) => {
                const settings = getSettings();
                toggle.setValue(settings.useFilenameAsScheduledDate).onChange(async (value) => {
                    updateSettings({ useFilenameAsScheduledDate: value });
                    setSettingVisibility(scheduledDateExtraFormat, value);
                    setSettingVisibility(scheduledDateFolders, value);
                    await this.plugin.saveSettings();
                });
            });

        scheduledDateExtraFormat = new Setting(containerEl)
            .setName(i18n.t('settings.datesFromFileNames.scheduledDate.extraFormat.name'))
            .setDesc(
                SettingsTab.createFragmentWithHTML(
                    i18n.t('settings.datesFromFileNames.scheduledDate.extraFormat.description.line1') +
                        '</br>' +
                        `<p>${i18n.t('settings.changeRequiresRestart')}</p>` +
                        `<p><a href="https://momentjs.com/docs/#/displaying/format/">${i18n.t(
                            'settings.datesFromFileNames.scheduledDate.extraFormat.description.line2',
                        )}</a></p>`,
                ),
            )
            .addText((text) => {
                const settings = getSettings();

                text.setPlaceholder(i18n.t('settings.datesFromFileNames.scheduledDate.extraFormat.placeholder'))
                    .setValue(settings.filenameAsScheduledDateFormat)
                    .onChange(async (value) => {
                        updateSettings({ filenameAsScheduledDateFormat: value });
                        await this.plugin.saveSettings();
                    });
            });

        scheduledDateFolders = new Setting(containerEl)
            .setName(i18n.t('settings.datesFromFileNames.scheduledDate.folders.name'))
            .setDesc(
                SettingsTab.createFragmentWithHTML(
                    `<p>${i18n.t('settings.datesFromFileNames.scheduledDate.folders.description')}</p>` +
                        `<p>${i18n.t('settings.changeRequiresRestart')}</p>`,
                ),
            )
            .addText(async (input) => {
                const settings = getSettings();
                await this.plugin.saveSettings();
                input
                    .setValue(SettingsTab.renderFolderArray(settings.filenameAsDateFolders))
                    .onChange(async (value) => {
                        const folders = SettingsTab.parseCommaSeparatedFolders(value);
                        updateSettings({ filenameAsDateFolders: folders });
                        await this.plugin.saveSettings();
                    });
            });
        setSettingVisibility(scheduledDateExtraFormat, getSettings().useFilenameAsScheduledDate);
        setSettingVisibility(scheduledDateFolders, getSettings().useFilenameAsScheduledDate);

        // ---------------------------------------------------------------------------
        new Setting(containerEl).setName(i18n.t('settings.recurringTasks.heading')).setHeading();
        // ---------------------------------------------------------------------------

        new Setting(containerEl)
            .setName(i18n.t('settings.recurringTasks.nextLine.name'))
            .setDesc(
                SettingsTab.createFragmentWithHTML(
                    i18n.t('settings.recurringTasks.nextLine.description') +
                        '</br>' +
                        this.seeTheDocumentation('https://publish.obsidian.md/tasks/Getting+Started/Recurring+Tasks'),
                ),
            )
            .addToggle((toggle) => {
                const { recurrenceOnNextLine: recurrenceOnNextLine } = getSettings();
                toggle.setValue(recurrenceOnNextLine).onChange(async (value) => {
                    updateSettings({ recurrenceOnNextLine: value });
                    await this.plugin.saveSettings();
                });
            });

        new Setting(containerEl)
            .setName(i18n.t('settings.recurringTasks.removeScheduledDate.name'))
            .setDesc(
                SettingsTab.createFragmentWithHTML(
                    i18n.t('settings.recurringTasks.removeScheduledDate.description.line1') +
                        '</br>' +
                        i18n.t('settings.recurringTasks.removeScheduledDate.description.line2') +
                        '</br>' +
                        this.seeTheDocumentation('https://publish.obsidian.md/tasks/Getting+Started/Recurring+Tasks'),
                ),
            )
            .addToggle((toggle) => {
                const { removeScheduledDateOnRecurrence } = getSettings();
                toggle.setValue(removeScheduledDateOnRecurrence).onChange(async (value) => {
                    updateSettings({ removeScheduledDateOnRecurrence: value });
                    await this.plugin.saveSettings();
                });
            });

        // ---------------------------------------------------------------------------
        new Setting(containerEl).setName(i18n.t('settings.autoSuggest.heading')).setHeading();
        // ---------------------------------------------------------------------------
        let autoSuggestMinimumMatchLength: Setting | null = null;
        let autoSuggestMaximumSuggestions: Setting | null = null;

        new Setting(containerEl)
            .setName(i18n.t('settings.autoSuggest.toggle.name'))
            .setDesc(
                SettingsTab.createFragmentWithHTML(
                    i18n.t('settings.autoSuggest.toggle.description') +
                        '</br>' +
                        `<p>${i18n.t('settings.changeRequiresRestart')}</p>` +
                        this.seeTheDocumentation('https://publish.obsidian.md/tasks/Getting+Started/Auto-Suggest'),
                ),
            )
            .addToggle((toggle) => {
                const settings = getSettings();
                toggle.setValue(settings.autoSuggestInEditor).onChange(async (value) => {
                    updateSettings({ autoSuggestInEditor: value });
                    await this.plugin.saveSettings();
                    setSettingVisibility(autoSuggestMinimumMatchLength, value);
                    setSettingVisibility(autoSuggestMaximumSuggestions, value);
                });
            });

        autoSuggestMinimumMatchLength = new Setting(containerEl)
            .setName(i18n.t('settings.autoSuggest.minLength.name'))
            .setDesc(
                SettingsTab.createFragmentWithHTML(
                    `<p>${i18n.t('settings.autoSuggest.minLength.description')}</p>` +
                        `<p>${i18n.t('settings.changeRequiresRestart')}</p>`,
                ),
            )
            .addSlider((slider) => {
                const settings = getSettings();
                slider
                    .setLimits(0, 3, 1)
                    .setValue(settings.autoSuggestMinMatch)
                    .setDynamicTooltip()
                    .onChange(async (value) => {
                        updateSettings({ autoSuggestMinMatch: value });
                        await this.plugin.saveSettings();
                    });
            });

        autoSuggestMaximumSuggestions = new Setting(containerEl)
            .setName(i18n.t('settings.autoSuggest.maxSuggestions.name'))
            .setDesc(
                SettingsTab.createFragmentWithHTML(
                    `<p>${i18n.t('settings.autoSuggest.maxSuggestions.description')}</p>` +
                        `<p>${i18n.t('settings.changeRequiresRestart')}</p>`,
                ),
            )
            .addSlider((slider) => {
                const settings = getSettings();
                slider
                    .setLimits(3, 20, 1)
                    .setValue(settings.autoSuggestMaxItems)
                    .setDynamicTooltip()
                    .onChange(async (value) => {
                        updateSettings({ autoSuggestMaxItems: value });
                        await this.plugin.saveSettings();
                    });
            });
        setSettingVisibility(autoSuggestMinimumMatchLength, getSettings().autoSuggestInEditor);
        setSettingVisibility(autoSuggestMaximumSuggestions, getSettings().autoSuggestInEditor);

        // ---------------------------------------------------------------------------
        new Setting(containerEl).setName(i18n.t('settings.dialogs.heading')).setHeading();
        // ---------------------------------------------------------------------------

        new Setting(containerEl)
            .setName(i18n.t('settings.dialogs.accessKeys.name'))
            .setDesc(
                SettingsTab.createFragmentWithHTML(
                    i18n.t('settings.dialogs.accessKeys.description') +
                        '</br>' +
                        this.seeTheDocumentation(
                            'https://publish.obsidian.md/tasks/Getting+Started/Create+or+edit+Task#Keyboard+shortcuts',
                        ),
                ),
            )
            .addToggle((toggle) => {
                const settings = getSettings();
                toggle.setValue(settings.provideAccessKeys).onChange(async (value) => {
                    updateSettings({ provideAccessKeys: value });
                    await this.plugin.saveSettings();
                });
            });

        // ---------------------------------------------------------------------------
        new Setting(containerEl).setName('Microsoft Planner Integration').setHeading();
        // ---------------------------------------------------------------------------

        this.renderPlannerSettings(containerEl);
    }

    // ---------------------------------------------------------------------------
    // Microsoft Planner settings section
    // ---------------------------------------------------------------------------

    private renderPlannerSettings(containerEl: HTMLElement): void {
        const ps = getSettings().plannerSettings;

        // -- Enable/disable toggle --
        new Setting(containerEl)
            .setName('Enable Planner sync')
            .setDesc('Sync Obsidian tasks with Microsoft Planner.')
            .addToggle((toggle) => {
                toggle.setValue(ps.enabled).onChange(async (value) => {
                    updateSettings({ plannerSettings: { ...getSettings().plannerSettings, enabled: value } });
                    await this.plugin.saveSettings();
                    // Refresh to show/hide remaining fields
                    this.display();
                });
            });

        if (!ps.enabled) return;

        // -- Azure AD credentials --
        const setupHint = containerEl.createEl('p', {});
        setupHint.style.cssText = 'font-size:0.85em; color:var(--text-muted); margin:0 0 12px; padding:8px 12px; border-left:3px solid var(--interactive-accent);';
        setupHint.innerHTML =
            '<strong>Azure AD app setup checklist:</strong><br>' +
            '1. Register an app in <a href="https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps" target="_blank">Azure Portal → App registrations</a><br>' +
            '2. Under <strong>Authentication</strong> → <em>Advanced settings</em>: set <strong>"Allow public client flows" → Yes</strong><br>' +
            '3. Under <strong>API permissions</strong>: add <code>Tasks.ReadWrite</code> and <code>User.Read</code> (Microsoft Graph, Delegated)<br>' +
            '4. Copy the Tenant ID and Client ID from the app <strong>Overview</strong> page into the fields below.';

        new Setting(containerEl)
            .setName('Tenant ID')
            .setDesc(
                'The Directory (tenant) ID from your Azure AD app registration. ' +
                    'Found under Azure Portal → App registrations → your app → Overview.',
            )
            .addText((text) => {
                text.setPlaceholder('xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx')
                    .setValue(ps.tenantId)
                    .onChange(
                        debounce(async (value) => {
                            updateSettings({
                                plannerSettings: { ...getSettings().plannerSettings, tenantId: value.trim() },
                            });
                            await this.plugin.saveSettings();
                        }, 500, true),
                    );
            });

        new Setting(containerEl)
            .setName('Client (Application) ID')
            .setDesc('The Application (client) ID from your Azure AD app registration.')
            .addText((text) => {
                text.setPlaceholder('xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx')
                    .setValue(ps.clientId)
                    .onChange(
                        debounce(async (value) => {
                            updateSettings({
                                plannerSettings: { ...getSettings().plannerSettings, clientId: value.trim() },
                            });
                            await this.plugin.saveSettings();
                        }, 500, true),
                    );
            });

        // -- Authentication --
        const authStatusEl = containerEl.createEl('p', {
            cls: 'setting-item-description',
            text: ps.userDisplayName
                ? `Authenticated as: ${ps.userDisplayName}`
                : 'Not authenticated.',
        });

        const authSetting = new Setting(containerEl).setName('Authentication').setDesc('');

        let cancelBtnEl: HTMLElement | null = null;

        authSetting.addButton((btn) => {
            btn.setButtonText(ps.userDisplayName ? 'Re-authenticate' : 'Authenticate')
                .setCta()
                .onClick(async () => {
                    // Disable auth button and add Cancel inline — no full re-render
                    btn.setDisabled(true);
                    btn.setButtonText('Authenticating…');

                    if (!cancelBtnEl) {
                        authSetting.addButton((cancelBtn) => {
                            cancelBtn.setButtonText('Cancel').setWarning().onClick(() => {
                                this.authAbortController?.abort();
                                this.authAbortController = null;
                                authStatusEl.textContent = 'Authentication cancelled.';
                                btn.setDisabled(false);
                                btn.setButtonText(ps.userDisplayName ? 'Re-authenticate' : 'Authenticate');
                                cancelBtnEl?.remove();
                                cancelBtnEl = null;
                            });
                            cancelBtnEl = cancelBtn.buttonEl;
                        });
                    }

                    await this.startPlannerAuth(authStatusEl, () => {
                        // cleanup: remove cancel button, re-enable auth button
                        cancelBtnEl?.remove();
                        cancelBtnEl = null;
                    });
                });
        });

        if (!PlannerAuth.isAuthenticated(ps)) return;

        // -- Default plan & bucket --
        new Setting(containerEl)
            .setName('Default plan')
            .setDesc('Tasks with no matching tag are added to this plan. Usually "Private Tasks" for personal tasks.');

        this.renderPlanDropdown(containerEl, 'Default plan', ps.defaultPlanId, async (plan) => {
            updateSettings({
                plannerSettings: {
                    ...getSettings().plannerSettings,
                    defaultPlanId: plan.id,
                    defaultPlanTitle: plan.title,
                    defaultBucketId: '',
                    defaultBucketTitle: '',
                },
            });
            await this.plugin.saveSettings();
            this.display();
        });

        if (ps.defaultPlanId) {
            this.renderBucketDropdown(
                containerEl,
                'Default bucket',
                ps.defaultPlanId,
                ps.defaultBucketId,
                async (bucket) => {
                    updateSettings({
                        plannerSettings: {
                            ...getSettings().plannerSettings,
                            defaultBucketId: bucket.id,
                            defaultBucketTitle: bucket.name,
                        },
                    });
                    await this.plugin.saveSettings();
                    this.display(); // refresh so watched-buckets list excludes the new default
                },
            );

            this.renderWatchedBucketsSection(containerEl, ps.defaultPlanId);
        }

        // -- Tag → bucket mappings --
        new Setting(containerEl).setName('Tag to bucket mappings').setHeading();

        new Setting(containerEl)
            .setName('')
            .setDesc(
                'By default every task goes to the default plan/bucket. ' +
                    'Add a #tag to route it to a different plan or bucket instead. ' +
                    'The first matching tag wins.',
            );

        this.renderTagMappingRows(containerEl, ps.tagMappings, ps.defaultPlanId, ps.defaultPlanTitle);

        new Setting(containerEl).addButton((btn) => {
            btn.setButtonText('+ Add mapping').onClick(async () => {
                const current = getSettings().plannerSettings;
                updateSettings({
                    plannerSettings: {
                        ...current,
                        tagMappings: [
                            ...current.tagMappings,
                            {
                                tag: '',
                                planId: current.defaultPlanId,
                                planTitle: current.defaultPlanTitle,
                                bucketId: '',
                                bucketTitle: '',
                            },
                        ],
                    },
                });
                await this.plugin.saveSettings();
                this.display();
            });
        });

        // -- Priority mapping --
        new Setting(containerEl).setName('Priority mapping').setHeading();
        new Setting(containerEl)
            .setName('')
            .setDesc(
                'Map each Obsidian priority to a Planner priority. ' +
                    'Planner values: 0 = Urgent, 1 = Important, 5 = Medium, 9 = Low.',
            );

        const priorityKeys: Array<{ label: string; key: keyof typeof ps.priorityMapping }> = [
            { label: 'Highest 🔺', key: 'highest' },
            { label: 'High ⏫', key: 'high' },
            { label: 'Medium 🔼', key: 'medium' },
            { label: 'None (normal)', key: 'none' },
            { label: 'Low 🔽', key: 'low' },
            { label: 'Lowest ⏬', key: 'lowest' },
        ];

        for (const { label, key } of priorityKeys) {
            new Setting(containerEl).setName(label).addDropdown((dd) => {
                dd.addOption('0', '0 — Urgent');
                dd.addOption('1', '1 — Important');
                dd.addOption('5', '5 — Medium');
                dd.addOption('9', '9 — Low');
                dd.setValue(String(ps.priorityMapping[key]));
                dd.onChange(async (value) => {
                    const current = getSettings().plannerSettings;
                    updateSettings({
                        plannerSettings: {
                            ...current,
                            priorityMapping: {
                                ...current.priorityMapping,
                                [key]: Number(value) as PlannerPriority,
                            },
                        },
                    });
                    await this.plugin.saveSettings();
                });
            });
        }

        // -- Sync behaviour --
        new Setting(containerEl).setName('Sync behaviour').setHeading();

        new Setting(containerEl)
            .setName('Auto-assign task IDs')
            .setDesc(
                'Automatically add a 🆔 identifier to new tasks so they can be stably tracked ' +
                    'and synced with Planner.',
            )
            .addToggle((toggle) => {
                toggle.setValue(ps.autoAssignTaskIds).onChange(async (value) => {
                    updateSettings({
                        plannerSettings: { ...getSettings().plannerSettings, autoAssignTaskIds: value },
                    });
                    await this.plugin.saveSettings();
                });
            });

        new Setting(containerEl)
            .setName('Sync on file open')
            .setDesc('Pull Planner task changes when you open a file that contains linked tasks.')
            .addToggle((toggle) => {
                toggle.setValue(ps.syncOnFileOpen).onChange(async (value) => {
                    updateSettings({
                        plannerSettings: { ...getSettings().plannerSettings, syncOnFileOpen: value },
                    });
                    await this.plugin.saveSettings();
                });
            });

        new Setting(containerEl)
            .setName('Background sync interval (minutes)')
            .setDesc('Pull changes from Planner while the active file is open. Set to 0 to disable.')
            .addSlider((slider) => {
                slider
                    .setLimits(0, 60, 1)
                    .setValue(ps.syncIntervalMinutes)
                    .setDynamicTooltip()
                    .onChange(async (value) => {
                        updateSettings({
                            plannerSettings: { ...getSettings().plannerSettings, syncIntervalMinutes: value },
                        });
                        await this.plugin.saveSettings();
                        this.plannerSyncManager?.startSyncInterval();
                    });
            });

        // -------------------------------------------------------------------
        // Import wizard
        // -------------------------------------------------------------------

        containerEl.createEl('h3', { text: 'Import from Planner' });

        const importDesc = containerEl.createEl('p', {});
        importDesc.style.cssText = 'font-size:0.85em; color:var(--text-muted); margin:0 0 12px;';
        importDesc.textContent =
            'Fetch active tasks from a Planner plan and append them to a vault file. Already-linked tasks are skipped.';

        // State
        let importPlanId = '';
        let importBucketId = '';
        let importBuckets: PlannerBucket[] = [];
        let importTargetPath = 'Planner Tasks.md';

        // Result status paragraph — declared early so async callbacks can reference it
        const importStatusEl = containerEl.createEl('p', {});
        importStatusEl.style.cssText = 'font-size:0.85em; color:var(--text-muted); margin:4px 0 0; min-height:1.2em;';

        // Plan dropdown (async-loaded)
        new Setting(containerEl)
            .setName('Plan')
            .setDesc('Choose the Planner plan to import tasks from.')
            .addDropdown((dd) => {
                dd.addOption('', 'Loading plans…');
                dd.setDisabled(true);

                if (!this.plannerSyncManager) {
                    dd.selectEl.empty();
                    dd.addOption('', '— Planner not initialised —');
                    return;
                }

                const client = this.plannerSyncManager.getApiClient();
                client
                    .getMyPlans()
                    .then((plans: PlannerPlan[]) => {
                        dd.selectEl.empty();
                        dd.addOption('', '— Select a plan —');
                        for (const p of plans) {
                            dd.addOption(p.id, p.title);
                        }
                        dd.setDisabled(false);
                        dd.onChange(async (planId) => {
                            importPlanId = planId;
                            importBucketId = '';
                            importBuckets = [];
                            bucketContainer.empty();
                            if (!planId) return;
                            try {
                                importBuckets = await client.getPlanBuckets(planId);
                                renderBucketDropdown();
                            } catch {
                                importStatusEl.textContent = 'Failed to load buckets.';
                            }
                        });
                    })
                    .catch(() => {
                        dd.selectEl.empty();
                        dd.addOption('', '— Failed to load plans —');
                    });
            });

        // Bucket container — inserted right after Plan in the DOM so it appears below it
        const bucketContainer = containerEl.createDiv();

        const renderBucketDropdown = (): void => {
            bucketContainer.empty();
            if (!importPlanId || importBuckets.length === 0) return;

            new Setting(bucketContainer)
                .setName('Bucket (optional)')
                .setDesc('Limit import to this bucket. Leave at "All buckets" to import everything.')
                .addDropdown((dd) => {
                    dd.addOption('', 'All buckets');
                    for (const b of importBuckets) {
                        dd.addOption(b.id, b.name);
                    }
                    dd.setValue(importBucketId);
                    dd.onChange((v) => {
                        importBucketId = v;
                    });
                });
        };

        // Target file path input
        new Setting(containerEl)
            .setName('Target file')
            .setDesc('Vault-relative path of the file to append imported tasks to (created if absent).')
            .addText((text) => {
                text.setPlaceholder('Planner Tasks.md')
                    .setValue(importTargetPath)
                    .onChange((v) => {
                        importTargetPath = v.trim() || 'Planner Tasks.md';
                    });
            })
            .addButton((btn) => {
                btn.setButtonText('Import').onClick(async () => {
                    if (!importPlanId) {
                        importStatusEl.textContent = 'Please select a plan first.';
                        return;
                    }
                    if (!this.plannerSyncManager) {
                        importStatusEl.textContent = 'Planner sync manager not available.';
                        return;
                    }
                    btn.setDisabled(true);
                    btn.setButtonText('Importing…');
                    importStatusEl.textContent = 'Importing…';
                    try {
                        const count = await this.plannerSyncManager.importFromPlanner(
                            importPlanId,
                            importBucketId || undefined,
                            importTargetPath,
                        );
                        importStatusEl.textContent =
                            count === 0
                                ? 'No new tasks to import (all already linked or plan is empty).'
                                : `Imported ${count} task${count === 1 ? '' : 's'} → ${importTargetPath}`;
                    } catch (err) {
                        importStatusEl.textContent = `Import failed: ${err instanceof Error ? err.message : String(err)}`;
                    } finally {
                        btn.setDisabled(false);
                        btn.setButtonText('Import');
                    }
                });
            });
    }

    // -----------------------------------------------------------------------
    // Planner settings helpers
    // -----------------------------------------------------------------------

    private async startPlannerAuth(statusEl: HTMLElement, onDone: () => void): Promise<void> {
        const ps = getSettings().plannerSettings;
        if (!ps.tenantId || !ps.clientId) {
            new Notice('Please enter Tenant ID and Client ID before authenticating.');
            onDone();
            return;
        }

        this.authAbortController = new AbortController();

        try {
            const dc = await PlannerAuth.startDeviceCodeFlow(ps.tenantId, ps.clientId);
            statusEl.innerHTML =
                `<strong>Step 1:</strong> Open ` +
                `<a href="${dc.verification_uri}" target="_blank">${dc.verification_uri}</a> ` +
                `in your browser and enter code: <strong>${dc.user_code}</strong><br>` +
                `<em>Waiting for authorisation…</em>`;

            const tokens = await PlannerAuth.pollForToken(
                ps.tenantId,
                ps.clientId,
                dc.device_code,
                dc.interval,
                this.authAbortController.signal,
            );

            // Save tokens
            updateSettings({
                plannerSettings: {
                    ...getSettings().plannerSettings,
                    accessToken: tokens.access_token,
                    accessTokenExpiresAt: Date.now() + tokens.expires_in * 1000,
                    refreshToken: tokens.refresh_token ?? '',
                },
            });
            await this.plugin.saveSettings();

            // Fetch and save user info
            const client = this.plannerSyncManager?.getApiClient();
            if (client) {
                try {
                    const me = await client.getMe();
                    updateSettings({
                        plannerSettings: {
                            ...getSettings().plannerSettings,
                            userId: me.id,
                            userDisplayName: me.displayName,
                        },
                    });
                    await this.plugin.saveSettings();
                } catch {
                    // Non-fatal — user info can be fetched later
                }
            }

            this.authAbortController = null;
            onDone();
            new Notice('Planner: authentication successful!');
            this.display(); // full re-render only after completion
        } catch (err: unknown) {
            this.authAbortController = null;
            const raw = err instanceof Error ? err.message : String(err);
            const message = raw.includes('client_assertion or client_secret')
                ? 'App is not configured as a public client. ' +
                  'In Azure Portal → App registrations → your app → Authentication → ' +
                  'Advanced settings: set "Allow public client flows" to Yes, then save.'
                : raw;
            statusEl.textContent = `Authentication failed: ${message}`;
            onDone();
        }
    }

    private renderPlanDropdown(
        containerEl: HTMLElement,
        label: string,
        currentPlanId: string,
        onChange: (plan: PlannerPlan) => Promise<void>,
    ): void {
        if (!this.plannerSyncManager) return;
        const client = this.plannerSyncManager.getApiClient();

        const setting = new Setting(containerEl).setName(label).setDesc('Loading plans…');

        client
            .getMyPlans()
            .then((plans) => {
                setting.setDesc('');
                setting.addDropdown((dd) => {
                    dd.addOption('', '— select plan —');
                    for (const plan of plans) dd.addOption(plan.id, plan.title);
                    dd.setValue(currentPlanId);
                    dd.onChange(async (id) => {
                        const plan = plans.find((p) => p.id === id);
                        if (plan) await onChange(plan);
                    });
                });
            })
            .catch(() => setting.setDesc('Failed to load plans. Check your credentials.'));
    }

    private renderWatchedBucketsSection(containerEl: HTMLElement, planId: string): void {
        if (!this.plannerSyncManager) return;
        const client = this.plannerSyncManager.getApiClient();

        // Create a scoped wrapper div synchronously so it occupies the correct
        // position in the DOM before the async bucket fetch completes.
        const wrapper = containerEl.createDiv();

        const heading = new Setting(wrapper)
            .setName('Additional watched buckets')
            .setDesc('Loading buckets…');

        client
            .getPlanBuckets(planId)
            .then((buckets) => {
                heading.setDesc(
                    'Buckets included in pull-sync and import alongside the default bucket. ' +
                        'The default bucket is always watched.',
                );

                const ps = getSettings().plannerSettings;

                // Render one toggle per bucket (excluding the default bucket)
                for (const bucket of buckets) {
                    if (bucket.id === ps.defaultBucketId) continue; // always watched, no toggle needed

                    const isWatched = ps.watchedBucketIds.includes(bucket.id);

                    new Setting(wrapper)
                        .setName(bucket.name)
                        .setClass('planner-watched-bucket-row')
                        .addToggle((toggle) => {
                            toggle.setValue(isWatched).onChange(async (enabled) => {
                                const current = getSettings().plannerSettings;
                                const updated = enabled
                                    ? [...current.watchedBucketIds, bucket.id]
                                    : current.watchedBucketIds.filter((id) => id !== bucket.id);
                                updateSettings({
                                    plannerSettings: { ...current, watchedBucketIds: updated },
                                });
                                await this.plugin.saveSettings();
                            });
                        });
                }

                if (buckets.filter((b) => b.id !== ps.defaultBucketId).length === 0) {
                    heading.setDesc('No other buckets found in this plan.');
                }
            })
            .catch(() => heading.setDesc('Failed to load buckets.'));
    }

    private renderBucketDropdown(
        containerEl: HTMLElement,
        label: string,
        planId: string,
        currentBucketId: string,
        onChange: (bucket: PlannerBucket) => Promise<void>,
    ): void {
        if (!this.plannerSyncManager) return;
        const client = this.plannerSyncManager.getApiClient();

        const setting = new Setting(containerEl).setName(label).setDesc('Loading buckets…');

        client
            .getPlanBuckets(planId)
            .then((buckets) => {
                setting.setDesc('');
                setting.addDropdown((dd) => {
                    dd.addOption('', '— select bucket —');
                    for (const b of buckets) dd.addOption(b.id, b.name);
                    dd.setValue(currentBucketId);
                    dd.onChange(async (id) => {
                        const bucket = buckets.find((b) => b.id === id);
                        if (bucket) await onChange(bucket);
                    });
                });
            })
            .catch(() => setting.setDesc('Failed to load buckets.'));
    }

    private renderTagMappingRows(
        containerEl: HTMLElement,
        mappings: TagBucketMapping[],
        defaultPlanId: string,
        defaultPlanTitle: string,
    ): void {
        if (mappings.length === 0) return;

        // Column header row
        const header = containerEl.createDiv({
            attr: {
                style: 'display:grid; grid-template-columns:1fr 1fr 1fr auto; gap:8px; ' +
                       'padding:4px 16px; font-size:0.8em; color:var(--text-muted);',
            },
        });
        header.createSpan({ text: '#tag' });
        header.createSpan({ text: '→ plan' });
        header.createSpan({ text: '→ bucket' });
        header.createSpan();

        mappings.forEach((mapping, index) => {
            const resolvedPlanId = mapping.planId || defaultPlanId;

            // Bucket dropdown — built once, reloaded when plan changes
            let bucketSelectEl: HTMLSelectElement | null = null;

            const row = new Setting(containerEl)
                .addText((text) => {
                    text.setPlaceholder('#tag')
                        .setValue(mapping.tag)
                        .onChange(
                            debounce(async (value) => {
                                const current = getSettings().plannerSettings;
                                const updated = [...current.tagMappings];
                                updated[index] = {
                                    ...updated[index],
                                    tag: value.trim(),
                                    planId: updated[index].planId || defaultPlanId,
                                    planTitle: updated[index].planTitle || defaultPlanTitle,
                                };
                                updateSettings({ plannerSettings: { ...current, tagMappings: updated } });
                                await this.plugin.saveSettings();
                            }, 500, true),
                        );
                    text.inputEl.style.width = '110px';
                })
                .addDropdown((planDd) => {
                    // Show current plan immediately, then load full list async
                    planDd.addOption(resolvedPlanId, mapping.planTitle || defaultPlanTitle || resolvedPlanId);
                    planDd.setValue(resolvedPlanId);

                    if (!this.plannerSyncManager) return;
                    const client = this.plannerSyncManager.getApiClient();

                    client.getMyPlans().then((plans) => {
                        planDd.selectEl.empty();
                        for (const p of plans) planDd.addOption(p.id, p.title);
                        planDd.setValue(resolvedPlanId);

                        planDd.onChange(async (newPlanId) => {
                            const selectedPlan = plans.find((p) => p.id === newPlanId);
                            if (!selectedPlan) return;

                            // Persist plan change, clear bucket
                            const current = getSettings().plannerSettings;
                            const updated = [...current.tagMappings];
                            updated[index] = {
                                ...updated[index],
                                planId: selectedPlan.id,
                                planTitle: selectedPlan.title,
                                bucketId: '',
                                bucketTitle: '',
                            };
                            updateSettings({ plannerSettings: { ...current, tagMappings: updated } });
                            await this.plugin.saveSettings();

                            // Reload bucket dropdown for the new plan
                            if (bucketSelectEl) {
                                const bdEl = bucketSelectEl;
                                bdEl.empty();
                                const loadingOpt = document.createElement('option');
                                loadingOpt.value = '';
                                loadingOpt.textContent = 'Loading…';
                                bdEl.appendChild(loadingOpt);

                                client.getPlanBuckets(selectedPlan.id).then((buckets) => {
                                    bdEl.empty();
                                    const placeholder = document.createElement('option');
                                    placeholder.value = '';
                                    placeholder.textContent = '— select bucket —';
                                    bdEl.appendChild(placeholder);
                                    for (const b of buckets) {
                                        const o = document.createElement('option');
                                        o.value = b.id;
                                        o.textContent = b.name;
                                        bdEl.appendChild(o);
                                    }
                                    bdEl.value = '';
                                }).catch(() => {/* silent */});
                            }
                        });
                    }).catch(() => {/* silent */});
                })
                .addDropdown((dd) => {
                    bucketSelectEl = dd.selectEl;
                    dd.addOption('', mapping.bucketTitle || '— bucket —');

                    if (resolvedPlanId) {
                        this.plannerSyncManager
                            ?.getApiClient()
                            .getPlanBuckets(resolvedPlanId)
                            .then((buckets) => {
                                dd.selectEl.empty();
                                dd.addOption('', '— select bucket —');
                                for (const b of buckets) dd.addOption(b.id, b.name);
                                dd.setValue(mapping.bucketId);
                                dd.onChange(async (id) => {
                                    const bucket = buckets.find((b) => b.id === id);
                                    if (!bucket) return;
                                    const current = getSettings().plannerSettings;
                                    const updated = [...current.tagMappings];
                                    updated[index] = {
                                        ...updated[index],
                                        bucketId: bucket.id,
                                        bucketTitle: bucket.name,
                                    };
                                    updateSettings({ plannerSettings: { ...current, tagMappings: updated } });
                                    await this.plugin.saveSettings();
                                });
                            })
                            .catch(() => {/* silent */});
                    }
                })
                .addExtraButton((btn) => {
                    btn.setIcon('trash').setTooltip('Remove').onClick(async () => {
                        const current = getSettings().plannerSettings;
                        const updated = current.tagMappings.filter((_, i) => i !== index);
                        updateSettings({ plannerSettings: { ...current, tagMappings: updated } });
                        await this.plugin.saveSettings();
                        this.display();
                    });
                });
            row.nameEl.remove();
        });
    }

    private seeTheDocumentation(url: string) {
        return `<p><a href="${url}">${i18n.t('settings.seeTheDocumentation')}</a>.</p>`;
    }

    private addOneSettingsBlock(
        containerEl: HTMLElement,
        heading: any,
        headingOpened: HeadingState,
    ): HTMLDetailsElement {
        const detailsContainer = containerEl.createEl('details', {
            cls: 'tasks-nested-settings',
            attr: {
                ...(heading.open || headingOpened[heading.text] ? { open: true } : {}),
            },
        });
        detailsContainer.empty();
        detailsContainer.ontoggle = () => {
            headingOpened[heading.text] = detailsContainer.open;
            updateSettings({ headingOpened: headingOpened });
            this.plugin.saveSettings();
        };
        const summary = detailsContainer.createEl('summary');
        new Setting(summary).setHeading().setName(heading.text);
        summary.createDiv('collapser').createDiv('handle');

        // detailsContainer.createEl(heading.level as keyof HTMLElementTagNameMap, { text: heading.text });

        if (heading.notice !== null) {
            const notice = detailsContainer.createEl('div', {
                cls: heading.notice.class,
                text: heading.notice.text,
            });
            if (heading.notice.html !== null) {
                notice.insertAdjacentHTML('beforeend', heading.notice.html);
            }
        }

        // This will process all the settings from settingsConfiguration.json and render
        // them out reducing the duplication of the code in this file. This will become
        // more important as features are being added over time.
        heading.settings.forEach((setting: any) => {
            if (setting.featureFlag !== '' && !isFeatureEnabled(setting.featureFlag)) {
                // The settings configuration has a featureFlag set and the user has not
                // enabled it. Skip adding the settings option.
                return;
            }
            if (setting.type === 'checkbox') {
                new Setting(detailsContainer)
                    .setName(setting.name)
                    .setDesc(setting.description)
                    .addToggle((toggle) => {
                        const settings = getSettings();
                        if (!settings.generalSettings[setting.settingName]) {
                            updateGeneralSetting(setting.settingName, setting.initialValue);
                        }
                        toggle
                            .setValue(<boolean>settings.generalSettings[setting.settingName])
                            .onChange(async (value) => {
                                updateGeneralSetting(setting.settingName, value);
                                await this.plugin.saveSettings();
                            });
                    });
            } else if (setting.type === 'text') {
                new Setting(detailsContainer)
                    .setName(setting.name)
                    .setDesc(setting.description)
                    .addText((text) => {
                        const settings = getSettings();
                        if (!settings.generalSettings[setting.settingName]) {
                            updateGeneralSetting(setting.settingName, setting.initialValue);
                        }

                        const onChange = async (value: string) => {
                            updateGeneralSetting(setting.settingName, value);
                            await this.plugin.saveSettings();
                        };

                        text.setPlaceholder(setting.placeholder.toString())
                            .setValue(settings.generalSettings[setting.settingName].toString())
                            .onChange(debounce(onChange, 500, true));
                    });
            } else if (setting.type === 'textarea') {
                new Setting(detailsContainer)
                    .setName(setting.name)
                    .setDesc(setting.description)
                    .addTextArea((text) => {
                        const settings = getSettings();
                        if (!settings.generalSettings[setting.settingName]) {
                            updateGeneralSetting(setting.settingName, setting.initialValue);
                        }

                        const onChange = async (value: string) => {
                            updateGeneralSetting(setting.settingName, value);
                            await this.plugin.saveSettings();
                        };

                        text.setPlaceholder(setting.placeholder.toString())
                            .setValue(settings.generalSettings[setting.settingName].toString())
                            .onChange(debounce(onChange, 500, true));

                        text.inputEl.rows = 8;
                        text.inputEl.cols = 40;
                    });
            } else if (setting.type === 'function') {
                this.customFunctions[setting.settingName](detailsContainer, this);
            }

            if (setting.notice !== null) {
                const notice = detailsContainer.createEl('p', {
                    cls: setting.notice.class,
                    text: setting.notice.text,
                });
                if (setting.notice.html !== null) {
                    notice.insertAdjacentHTML('beforeend', setting.notice.html);
                }
            }
        });

        return detailsContainer;
    }

    private static parseCommaSeparatedFolders(input: string): string[] {
        return (
            input
                // a limitation is that folder names may not contain commas
                .split(',')
                .map((folder) => folder.trim())
                // remove leading and trailing slashes
                .map((folder) => folder.replace(/^\/|\/$/g, ''))
                .filter((folder) => folder !== '')
        );
    }
    private static renderFolderArray(folders: string[]): string {
        return folders.join(',');
    }

    /**
     * Settings for Core Task Status
     * These are built-in statuses that can have minimal edits made,
     * but are not allowed to be deleted or added to.
     *
     * @param {HTMLElement} containerEl
     * @param {SettingsTab} settings
     */
    insertTaskCoreStatusSettings(containerEl: HTMLElement, settings: SettingsTab) {
        const { statusSettings } = getSettings();

        /* -------------------- One row per core status in the settings -------------------- */
        statusSettings.coreStatuses.forEach((status_type) => {
            createRowForTaskStatus(
                containerEl,
                status_type,
                statusSettings.coreStatuses,
                statusSettings,
                settings,
                settings.plugin,
                true, // isCoreStatus
            );
        });

        /* -------------------- 'Review and check your Statuses' button -------------------- */
        const createMermaidDiagram = new Setting(containerEl).addButton((button) => {
            const buttonName = i18n.t('settings.statuses.coreStatuses.buttons.checkStatuses.name');
            button
                .setButtonText(buttonName)
                .setCta()
                .onClick(async () => {
                    // Generate a new file unique file name, in the root of the vault
                    const now = window.moment();
                    const formattedDateTime = now.format('YYYY-MM-DD HH-mm-ss');
                    const filename = `Tasks Plugin - ${buttonName} ${formattedDateTime}.md`;

                    // Create the report
                    const version = this.plugin.manifest.version;
                    const statusRegistry = StatusRegistry.getInstance();
                    const fileContent = createStatusRegistryReport(statusSettings, statusRegistry, buttonName, version);

                    // Save the file
                    const file = await this.app.vault.create(filename, fileContent);

                    // And open the new file
                    const leaf = this.app.workspace.getLeaf(true);
                    await leaf.openFile(file);
                });
            button.setTooltip(i18n.t('settings.statuses.coreStatuses.buttons.checkStatuses.tooltip'));
        });
        createMermaidDiagram.infoEl.remove();
    }

    /**
     * Settings for Custom Task Status
     *
     * @param {HTMLElement} containerEl
     * @param {SettingsTab} settings
     */
    insertCustomTaskStatusSettings(containerEl: HTMLElement, settings: SettingsTab) {
        const { statusSettings } = getSettings();

        /* -------------------- One row per custom status in the settings -------------------- */
        statusSettings.customStatuses.forEach((status_type) => {
            createRowForTaskStatus(
                containerEl,
                status_type,
                statusSettings.customStatuses,
                statusSettings,
                settings,
                settings.plugin,
                false, // isCoreStatus
            );
        });

        containerEl.createEl('div');

        /* -------------------- 'Add New Task Status' button -------------------- */
        const setting = new Setting(containerEl).addButton((button) => {
            button
                .setButtonText(i18n.t('settings.statuses.customStatuses.buttons.addNewStatus.name'))
                .setCta()
                .onClick(async () => {
                    StatusSettings.addStatus(
                        statusSettings.customStatuses,
                        new StatusConfiguration('', '', '', false, StatusType.TODO),
                    );
                    await updateAndSaveStatusSettings(statusSettings, settings);
                });
        });
        setting.infoEl.remove();

        /* -------------------- Add all Status types supported by ... buttons -------------------- */
        type NamedTheme = [string, StatusCollection];
        const themes: NamedTheme[] = [
            // Light and Dark themes - alphabetical order
            [i18n.t('settings.statuses.collections.anuppuccinTheme'), Themes.anuppuccinSupportedStatuses()],
            [i18n.t('settings.statuses.collections.auraTheme'), Themes.auraSupportedStatuses()],
            [i18n.t('settings.statuses.collections.borderTheme'), Themes.borderSupportedStatuses()],
            [i18n.t('settings.statuses.collections.ebullientworksTheme'), Themes.ebullientworksSupportedStatuses()],
            [i18n.t('settings.statuses.collections.itsThemeAndSlrvbCheckboxes'), Themes.itsSupportedStatuses()],
            [i18n.t('settings.statuses.collections.minimalTheme'), Themes.minimalSupportedStatuses()],
            [i18n.t('settings.statuses.collections.thingsTheme'), Themes.thingsSupportedStatuses()],
            // Dark only themes - alphabetical order
            [i18n.t('settings.statuses.collections.lytModeTheme'), Themes.lytModeSupportedStatuses()],
        ];
        for (const [name, collection] of themes) {
            const addStatusesSupportedByThisTheme = new Setting(containerEl).addButton((button) => {
                const label = i18n.t('settings.statuses.collections.buttons.addCollection.name', {
                    themeName: name,
                    numberOfStatuses: collection.length,
                });
                button.setButtonText(label).onClick(async () => {
                    await addCustomStatesToSettings(collection, statusSettings, settings);
                });
            });
            addStatusesSupportedByThisTheme.infoEl.remove();
        }

        /* -------------------- 'Add All Unknown Status Types' button -------------------- */
        const addAllUnknownStatuses = new Setting(containerEl).addButton((button) => {
            button
                .setButtonText(i18n.t('settings.statuses.customStatuses.buttons.addAllUnknown.name'))
                .setCta()
                .onClick(async () => {
                    const tasks = this.plugin.getTasks();
                    const allStatuses = tasks!.map((task) => {
                        return task.status;
                    });
                    const unknownStatuses = StatusRegistry.getInstance().findUnknownStatuses(allStatuses);
                    if (unknownStatuses.length === 0) {
                        return;
                    }
                    unknownStatuses.forEach((s) => {
                        StatusSettings.addStatus(statusSettings.customStatuses, s);
                    });
                    await updateAndSaveStatusSettings(statusSettings, settings);
                });
        });
        addAllUnknownStatuses.infoEl.remove();

        /* -------------------- 'Reset Custom Status Types to Defaults' button -------------------- */
        const clearCustomStatuses = new Setting(containerEl).addButton((button) => {
            button
                .setButtonText(i18n.t('settings.statuses.customStatuses.buttons.resetCustomStatuses.name'))
                .setWarning()
                .onClick(async () => {
                    StatusSettings.resetAllCustomStatuses(statusSettings);
                    await updateAndSaveStatusSettings(statusSettings, settings);
                });
        });
        clearCustomStatuses.infoEl.remove();
    }
}

/**
 * Create the row to see and modify settings for a single task status type.
 * @param containerEl
 * @param statusType - The status type to be edited.
 * @param statuses - The list of statuses that statusType is stored in.
 * @param statusSettings - All the status types already in the user's settings, EXCEPT the standard ones.
 * @param settings
 * @param plugin
 * @param isCoreStatus - whether the status is a core status
 */
function createRowForTaskStatus(
    containerEl: HTMLElement,
    statusType: StatusConfiguration,
    statuses: StatusConfiguration[],
    statusSettings: StatusSettings,
    settings: SettingsTab,
    plugin: TasksPlugin,
    isCoreStatus: boolean,
) {
    //const taskStatusDiv = containerEl.createEl('div');

    const taskStatusPreview = containerEl.createEl('pre');
    taskStatusPreview.addClass('row-for-status');
    taskStatusPreview.textContent = new Status(statusType).previewText();

    const setting = new Setting(containerEl);

    setting.infoEl.replaceWith(taskStatusPreview);

    if (!isCoreStatus) {
        setting.addExtraButton((extra) => {
            extra
                .setIcon('cross')
                .setTooltip('Delete')
                .onClick(async () => {
                    if (StatusSettings.deleteStatus(statuses, statusType)) {
                        await updateAndSaveStatusSettings(statusSettings, settings);
                    }
                });
        });
    }

    setting.addExtraButton((extra) => {
        extra
            .setIcon('pencil')
            .setTooltip('Edit')
            .onClick(async () => {
                const modal = new CustomStatusModal(plugin, statusType, isCoreStatus);

                modal.onClose = async () => {
                    if (modal.saved) {
                        if (StatusSettings.replaceStatus(statuses, statusType, modal.statusConfiguration())) {
                            await updateAndSaveStatusSettings(statusSettings, settings);
                        }
                    }
                };

                modal.open();
            });
    });

    setting.infoEl.remove();
}

async function addCustomStatesToSettings(
    supportedStatuses: StatusCollection,
    statusSettings: StatusSettings,
    settings: SettingsTab,
) {
    const notices = StatusSettings.bulkAddStatusCollection(statusSettings, supportedStatuses);

    notices.forEach((notice) => {
        new Notice(notice);
    });

    await updateAndSaveStatusSettings(statusSettings, settings);
}

async function updateAndSaveStatusSettings(statusTypes: StatusSettings, settings: SettingsTab) {
    updateSettings({
        statusSettings: statusTypes,
    });

    // Update the active statuses.
    // This saves the user from having to restart Obsidian in order to apply the changed status(es).
    StatusSettings.applyToStatusRegistry(statusTypes, StatusRegistry.getInstance());

    await settings.saveSettings(true);
}

function makeMultilineTextSetting(setting: Setting) {
    const { settingEl, infoEl, controlEl } = setting;
    const textEl: HTMLElement | null = controlEl.querySelector('textarea');

    // Not a setting with a text field
    if (textEl === null) {
        return;
    }

    settingEl.style.display = 'block';
    infoEl.style.marginRight = '0px';
    textEl.style.minWidth = '-webkit-fill-available';
}

function setSettingVisibility(setting: Setting | null, visible: boolean) {
    if (setting) {
        // @ts-expect-error Setting.setVisibility() is not exposed in the API.
        // Source: https://discord.com/channels/686053708261228577/840286264964022302/1293725986042544139
        setting.setVisibility(visible);
    } else {
        console.warn('Setting has not be initialised. Can update visibility of setting UI - in setSettingVisibility');
    }
}
