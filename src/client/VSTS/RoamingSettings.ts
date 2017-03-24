import { IDropdownOption } from "office-ui-fabric-react";
import APTCache from "../stores/aptCache";

export class RoamingSettings {

    public static GetInstance(): RoamingSettings {
        if (!RoamingSettings.instance) {
            RoamingSettings.instance = new RoamingSettings();
        }
        return RoamingSettings.instance;
    }

    // singleton instance
    private static instance: RoamingSettings;

    // constants
    private static readonly VERSION: number = 1;
    private static readonly VERSION_KEY: string = "version";
    private static readonly ACCOUNT_KEY: string = "default_account";
    private static readonly PROJECT_KEY: string = "default_project";
    private static readonly TEAM_KEY: string = "default_team";
    private static readonly ACCOUNTS_KEY: string = "accounts";
    private static readonly PROJECTS_KEY: string = "projects";
    private static readonly TEAMS_KEY: string = "teams";
    private static readonly ID_KEY: string = "member_id";

    public isValid: boolean = false;
    public isFull: boolean = false;
    public account: string = "";
    public project: string = "";
    public team: string = "";
    public accounts: IDropdownOption[] = [];
    public projects: IDropdownOption[] = [];
    public teams: IDropdownOption[] = [];
    public id: string;

    private constructor() {
        let version: number = this.get<number>(RoamingSettings.VERSION_KEY);
        if (version === undefined || version < RoamingSettings.VERSION) {
            this.isValid = false;
        } else {
            this.isValid = true;
            this.account = this.get<string>(RoamingSettings.ACCOUNT_KEY) || "";
            this.project = this.get<string>(RoamingSettings.PROJECT_KEY) || "";
            this.team = this.get<string>(RoamingSettings.TEAM_KEY) || "";
            this.accounts = this.get<IDropdownOption[]>(RoamingSettings.ACCOUNTS_KEY) || [];
            this.projects = this.get<IDropdownOption[]>(RoamingSettings.PROJECTS_KEY) || [];
            this.teams = this.get<IDropdownOption[]>(RoamingSettings.TEAMS_KEY) || [];
            this.id = this.get<string>(RoamingSettings.ID_KEY);
        }
    }

    public updateFromCache(cache: APTCache): void {
        this.account = cache.account;
        this.accounts = cache.accounts;
        this.project = cache.project;
        this.projects = cache.projects;
        this.team = cache.team;
        this.teams = cache.teams;

        this.save();
    }

    public save(): void {
        this.set(RoamingSettings.VERSION_KEY, RoamingSettings.VERSION);
        this.set(RoamingSettings.ACCOUNT_KEY, this.account);
        this.set(RoamingSettings.PROJECT_KEY, this.project);
        this.set(RoamingSettings.TEAM_KEY, this.team);
        this.set(RoamingSettings.ACCOUNTS_KEY, this.accounts);
        this.set(RoamingSettings.PROJECTS_KEY, this.projects);
        this.set(RoamingSettings.TEAMS_KEY, this.teams);
        this.set(RoamingSettings.ID_KEY, this.id);
        Office.context.roamingSettings.saveAsync();
    }

    public clear(): void {
        this.remove(RoamingSettings.VERSION_KEY);
        this.remove(RoamingSettings.ACCOUNT_KEY);
        this.remove(RoamingSettings.PROJECT_KEY);
        this.remove(RoamingSettings.TEAM_KEY);
        this.remove(RoamingSettings.ACCOUNTS_KEY);
        this.remove(RoamingSettings.PROJECTS_KEY);
        this.remove(RoamingSettings.TEAMS_KEY);
        this.remove(RoamingSettings.ID_KEY);
        Office.context.roamingSettings.saveAsync();
        RoamingSettings.instance = null;
    }

    private get<T>(key: string): T {
        return Office.context.roamingSettings.get(key);
    }

    private set<T>(key: string, value: T): void {
        Office.context.roamingSettings.set(key, value);
    }

    private remove(key: string): void {
        Office.context.roamingSettings.remove(key);
    }
}