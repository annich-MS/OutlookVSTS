import { action, observable, computed } from "mobx";
import { IDropdownOption } from "office-ui-fabric-react";

import APTPopulateStage from "../models/aptPopulateStage";
import RoamingSettings from "../models/roamingSettings";
import { Rest, DropdownParseable } from "../utils/rest";

export default class APTCache {
    @computed get accounts(): IDropdownOption[] { return this._accounts; };
    @computed get projects(): IDropdownOption[] { return this._projects; };
    @computed get teams(): IDropdownOption[] { return this._teams; };
    @computed get account(): string { return this._account; };
    @computed get project(): string { return this._project; };
    @computed get team(): string { return this._team; };
    @computed get populateStage(): APTPopulateStage { return this._populateStage; };
    @computed get name(): string { return this._name; };

    public readonly id: number;

    @observable private _accounts: IDropdownOption[] = [];
    @observable private _projects: IDropdownOption[] = [];
    @observable private _teams: IDropdownOption[] = [];
    @observable private _account: string = "";
    @observable private _project: string = "";
    @observable private _team: string = "";
    @observable private _name: string = "";
    @observable private _populateStage: APTPopulateStage = APTPopulateStage.PrePopulate;

    public constructor() {
        this.id = Math.floor(Math.random() * 10000);
    }

    @action public async populate(): Promise<void> {
        await this.populateAccounts();
        await this.populateProjects();
        await this.populateTeams();
    }

    @action public async populateAccounts(): Promise<void> {
        this.setPopulateStage(APTPopulateStage.MidAccount);
        try {
            let rs: RoamingSettings = await RoamingSettings.GetInstance();
            let list: DropdownParseable[] = await Rest.getAccounts(rs.id);
            let output = this.convert(list, this._account);
            this.setAccounts(output.options, output.selected);
            this.setPopulateStage(APTPopulateStage.PostAccount);

        } catch (err) {
            throw err.toString("populate accounts");
        }
    }

    @action public setAccounts(accounts: IDropdownOption[], selected?: string): void {
        this._account = this.getSelected(accounts, selected);
        this._accounts = accounts;
    }

    @action public setAccount(account: string): void {
        this._account = account;
    }

    @action public async populateProjects(): Promise<void> {
        this.setPopulateStage(APTPopulateStage.MidProject);
        try {
            let list: DropdownParseable[] = await Rest.getProjects(this.account);
            let output = this.convert(list, this.project);
            this.setProjects(output.options, output.selected);
            this.setPopulateStage(APTPopulateStage.PostProject);
        } catch (error) {
            throw error.toString("populate projects");
        }
    }

    @action public setProjects(projects: IDropdownOption[], selected?: string): void {
        this._project = this.getSelected(projects, selected);
        this._projects = projects;
    }

    @action public setProject(project: string): void {
        this._project = project;
    }

    @action public async populateTeams(): Promise<void> {
        this.setPopulateStage(APTPopulateStage.MidTeam);
        try {
            let teams: DropdownParseable[] = await Rest.getTeams(this.project, this.account);
            let output = this.convert(teams, this.team);
            this.setTeams(output.options, output.selected);
            this.setPopulateStage(APTPopulateStage.PostPopulate);
        } catch (error) {
            throw error.toString("populate teams");
        }
        return;
    }

    @action public setTeams(teams: IDropdownOption[], selected?: string): void {
        this._team = this.getSelected(teams, selected);
        this._teams = teams;
    }

    @action public setTeam(team: string): void {
        this._team = team;
    }

    @action public setPopulateStage(stage: APTPopulateStage): void {
        this._populateStage = stage;
    }

    @action public setName(name: string): void {
        this._name = name;
    }

    @action public clear() {
        this._account = "";
        this._project = "";
        this._team = "";
        this._accounts = [];
        this._projects = [];
        this._teams = [];
        this._populateStage = APTPopulateStage.PrePopulate;
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

    private convert(list: DropdownParseable[], selected: string): { options: IDropdownOption[], selected: string } {
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

export const aptCache: APTCache = new APTCache();
