import { computed, observe } from "mobx";
import { IDropdownOption } from "office-ui-fabric-react";

import { DropdownParseable, Rest } from "../../rest";

import { RoamingSettings } from "../RoamingSettings";
import APTPopulateStage from "../../models/aptPopulateStage";
import APTCache from "../../stores/aptCache";

abstract class DropdownConfig {
    public static createAccountConfig(cache: APTCache): DropdownConfig {
        if (DropdownConfig.accountInstance === undefined || cache.id !== DropdownConfig.accountInstance._cache.id) {
            DropdownConfig.accountInstance = new AccountDropdownConfig(cache);
        }
        return DropdownConfig.accountInstance;
    };
    public static createProjectConfig(cache: APTCache): DropdownConfig {
        if (DropdownConfig.projectInstance === undefined || cache.id !== DropdownConfig.projectInstance._cache.id) {
            DropdownConfig.projectInstance = new ProjectDropdownConfig(cache);
        }
        return DropdownConfig.projectInstance;
    }
    public static createTeamConfig(cache: APTCache): DropdownConfig {
        if (DropdownConfig.teamInstance === undefined || cache.id !== DropdownConfig.teamInstance._cache.id) {
            DropdownConfig.teamInstance = new TeamDropdownConfig(cache);
        }
        return DropdownConfig.teamInstance;
    }

    private static accountInstance: AccountDropdownConfig;
    private static projectInstance: ProjectDropdownConfig;
    private static teamInstance: TeamDropdownConfig;

    public label: string;
    public handleFailure: (error: string) => void;
    public abstract get isDisabled(): boolean;
    public abstract get options(): IDropdownOption[];
    public abstract get selected(): string;

    protected _cache: APTCache;

    protected constructor(cache: APTCache) {
        this._cache = cache;
    }

    public abstract changeSelected(selected: IDropdownOption);
    public populateIfNeeded(): void { /* Default Empty Implementation */ }

    protected abstract populate(): void;

}

class AccountDropdownConfig extends DropdownConfig {
    public readonly label: string = "Accounts";
    @computed public get isDisabled(): boolean { return this._cache.populateStage < APTPopulateStage.PostAccount; };
    @computed public get options(): IDropdownOption[] { return this._cache.accounts; };
    @computed public get selected(): string { return this._cache.account; };

    public constructor(cache: APTCache) {
        super(cache);
    }

    public populateIfNeeded(): void {
        if (this._cache.populateStage === APTPopulateStage.PrePopulate) {
            this.populate();
        }
    }

    public changeSelected(selected: IDropdownOption) {
        this._cache.setAccount(selected.text);
    }

    protected populate(): void {
        this._cache.populateAccounts();
    }
}

class ProjectDropdownConfig extends DropdownConfig {
    public readonly label: string = "Projects";
    @computed public get isDisabled(): boolean { return this._cache.populateStage < APTPopulateStage.PostProject; };
    @computed public get options(): IDropdownOption[] { return this._cache.projects; };
    @computed public get selected(): string { return this._cache.project; };

    public constructor(cache: APTCache) {
        super(cache);
        observe(this._cache, (change) => { if (change.name === "_account") { this.populate(); } });
    }

    public changeSelected(selected: IDropdownOption) {
        this._cache.setProject(selected.text);
    }

    protected populate(): void {
        this._cache.populateProjects();
    }
}

class TeamDropdownConfig extends DropdownConfig {
    public readonly label: string = "Teams";
    @computed public get isDisabled(): boolean { return this._cache.populateStage < APTPopulateStage.PostPopulate; };
    @computed public get options(): IDropdownOption[] { return this._cache.teams; };
    @computed public get selected(): string { return this._cache.team; };

    public constructor(cache: APTCache) {
        super(cache);
        observe(this._cache, (change) => { if (change.name === "_project") { this.populate(); } });
    }

    public changeSelected(selected: IDropdownOption) {
        this._cache.setTeam(selected.text);
    }

    protected populate(): void {
        this._cache.populateTeams();
    }
}

export default DropdownConfig;
