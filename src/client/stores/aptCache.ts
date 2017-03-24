import { action, observable, computed } from "mobx";
import { IDropdownOption } from "office-ui-fabric-react";

import APTPopulateStage from "../models/aptPopulateStage";

export default class APTCache {
    @computed get accounts(): IDropdownOption[] { return this._accounts; };
    @computed get projects(): IDropdownOption[] { return this._projects; };
    @computed get teams(): IDropdownOption[] { return this._teams; };
    @computed get account(): string { return this._account; };
    @computed get project(): string { return this._project; };
    @computed get team(): string { return this._team; };
    @computed get populateStage(): APTPopulateStage { return this._populateStage; };

    public readonly id: number;

    @observable private _accounts: IDropdownOption[] = [];
    @observable private _projects: IDropdownOption[] = [];
    @observable private _teams: IDropdownOption[] = [];
    @observable private _account: string = "";
    @observable private _project: string = "";
    @observable private _team: string = "";
    @observable private _populateStage: APTPopulateStage = APTPopulateStage.PostPopulate;

    public constructor() {
        this.id = Math.floor(Math.random() * 10000);
    }

    @action public setAccounts(accounts: IDropdownOption[], selected?: string): void {
        this._account = this.getSelected(accounts, selected);
        this._accounts = accounts;
    }

    @action public setProjects(projects: IDropdownOption[], selected?: string): void {
        this._project = this.getSelected(projects, selected);
        this._projects = projects;
    }

    @action public setTeams(teams: IDropdownOption[], selected?: string): void {
        this._team = this.getSelected(teams, selected);
        this._teams = teams;
    }

    @action public setAccount(account: string): void {
        this._account = account;
    }

    @action public setProject(project: string): void {
        this._project = project;
    }

    @action public setTeam(team: string): void {
        this._team = team;
    }

    @action public setPopulateStage(stage: APTPopulateStage) {
        this._populateStage = stage;
    }

    private getSelected(array: IDropdownOption[], selected: string): string {
        let retSelected: string = selected;
        array.forEach((value: IDropdownOption, index: number) => {
            if (selected === value.text || index === 0) {
                retSelected = value.text;
            }
        });
        return retSelected;
    }
}

export const aptCache: APTCache = new APTCache();
