import { computed, observe } from "mobx";
import { IDropdownOption } from "office-ui-fabric-react";

import { DropdownParseable, Rest, RestError } from "../../rest";

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

    protected abstract populate(): void;
    protected convert(list: DropdownParseable[], selected: string): { options: IDropdownOption[], selected: string } {
        list = list.sort(DropdownParseable.compare);
        let retSelected: string = selected;
        // see if selected exists
        return {
            options: list.map((element, index) => {
                let retVal: IDropdownOption = {
                    key: element.name,
                    text: element.name,
                };
                if (element.name === selected || index === 0) {
                    retSelected = element.name;
                }
                return retVal;
            }),
            selected: retSelected,
        };
    }
}

class AccountDropdownConfig extends DropdownConfig {
    public readonly label: string = "Accounts";
    @computed public get isDisabled(): boolean { return this._cache.populateStage < APTPopulateStage.PostAccount; };
    @computed public get options(): IDropdownOption[] { return this._cache.accounts; };
    @computed public get selected(): string { return this._cache.account; };

    public constructor(cache: APTCache) {
        super(cache);
    }

    public changeSelected(selected: IDropdownOption) {
        this._cache.setAccount(selected.text);
    }

    protected populate(): void {
        this._cache.setPopulateStage(APTPopulateStage.MidAccount);
        Rest.getAccounts(RoamingSettings.GetInstance().id, (error, list) => {
            if (error) {
                this.handleFailure(error.toString("populate accounts"));
                return;
            }
            let output = this.convert(list, this._cache.account);
            this._cache.setAccounts(output.options, output.selected);
            this._cache.setPopulateStage(APTPopulateStage.PostAccount);
        });
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
        this._cache.setPopulateStage(APTPopulateStage.MidProject);
        Rest.getProjects(this._cache.account, (error, list) => {
            if (error) {
                this.handleFailure(error.toString("populate projects"));
                return;
            }
            let output = this.convert(list, this._cache.project);
            this._cache.setProjects(output.options, output.selected);
            this._cache.setPopulateStage(APTPopulateStage.PostProject);
        });
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
        this._cache.setPopulateStage(APTPopulateStage.MidTeam);
        Rest.getTeams(this._cache.project, this._cache.account, (error, list) => {
            if (error) {
                this.handleFailure(error.toString("populate teams"));
                return;
            }
            let output = this.convert(list, this._cache.team);
            this._cache.setTeams(output.options, output.selected);
            this._cache.setPopulateStage(APTPopulateStage.PostPopulate);
        });
    }
}

export default DropdownConfig;

