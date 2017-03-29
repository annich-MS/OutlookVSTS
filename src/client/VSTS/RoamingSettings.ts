import { IDropdownOption } from "office-ui-fabric-react";
import IVSTSConfig from "../models/vstsConfig";
import APTCache from "../stores/aptCache";

abstract class BaseRoamingSettings {
    protected static readonly VERSION_KEY: string = "version";

    public abstract save(): Promise<void>;
    public abstract clear(): Promise<void>;

    protected get<T>(key: string): T {
        return Office.context.roamingSettings.get(key);
    }

    protected set<T>(key: string, value: T): void {
        Office.context.roamingSettings.set(key, value);
    }

    protected remove(key: string): void {
        Office.context.roamingSettings.remove(key);
    }

}

class RoamingSettings01 extends BaseRoamingSettings {

    // constants
    public static readonly VERSION: number = 1;
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

    public constructor() {
        super();
        let version: number = this.get<number>(RoamingSettings.VERSION_KEY);
        if (version === undefined || version < RoamingSettings01.VERSION) {
            this.isValid = false;
        } else {
            this.isValid = true;
            this.account = this.get<string>(RoamingSettings01.ACCOUNT_KEY) || "";
            this.project = this.get<string>(RoamingSettings01.PROJECT_KEY) || "";
            this.team = this.get<string>(RoamingSettings01.TEAM_KEY) || "";
            this.accounts = this.get<IDropdownOption[]>(RoamingSettings01.ACCOUNTS_KEY) || [];
            this.projects = this.get<IDropdownOption[]>(RoamingSettings01.PROJECTS_KEY) || [];
            this.teams = this.get<IDropdownOption[]>(RoamingSettings01.TEAMS_KEY) || [];
            this.id = this.get<string>(RoamingSettings01.ID_KEY);
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

    public save(): Promise<void> {
        return new Promise<void>((resolve, reject) => {
            this.set(RoamingSettings.VERSION_KEY, RoamingSettings01.VERSION);
            this.set(RoamingSettings01.ACCOUNT_KEY, this.account);
            this.set(RoamingSettings01.PROJECT_KEY, this.project);
            this.set(RoamingSettings01.TEAM_KEY, this.team);
            this.set(RoamingSettings01.ACCOUNTS_KEY, this.accounts);
            this.set(RoamingSettings01.PROJECTS_KEY, this.projects);
            this.set(RoamingSettings01.TEAMS_KEY, this.teams);
            this.set(RoamingSettings01.ID_KEY, this.id);
            Office.context.roamingSettings.saveAsync((result: Office.AsyncResult) => {
                if (result.error) {
                    reject(result.error);
                } else {
                    resolve();
                }
            });
        });
    }

    public clear(): Promise<void> {
        return new Promise<void>((resolve, reject) => {

            this.remove(RoamingSettings.VERSION_KEY);
            this.remove(RoamingSettings01.ACCOUNT_KEY);
            this.remove(RoamingSettings01.PROJECT_KEY);
            this.remove(RoamingSettings01.TEAM_KEY);
            this.remove(RoamingSettings01.ACCOUNTS_KEY);
            this.remove(RoamingSettings01.PROJECTS_KEY);
            this.remove(RoamingSettings01.TEAMS_KEY);
            this.remove(RoamingSettings01.ID_KEY);
            Office.context.roamingSettings.saveAsync((result: Office.AsyncResult) => {
                if (result.error) {
                    reject(result.error);
                } else {
                    RoamingSettings.ResetInstance();
                    resolve();
                }
            });
        });
    }
}

class RoamingSettings02 extends BaseRoamingSettings {

    public static readonly VERSION: number = 2;

    private static readonly CONFIGS_KEY: string = "configs";

    public configs: IVSTSConfig[];

    public constructor() {
        super();
        let version: number = this.get<number>(RoamingSettings.VERSION_KEY);
        switch (version) {
            case RoamingSettings01.VERSION:
                this.migrateFromRS01();
            case RoamingSettings02.VERSION:
                this.populate();
            default:
                this.preload();
        }

    }

    public save(): Promise<void> {
        return new Promise<void>((resolve, reject) => {
            this.set<IVSTSConfig[]>(RoamingSettings02.CONFIGS_KEY, this.configs);
            Office.context.roamingSettings.saveAsync((result: Office.AsyncResult) => {
                if (result.error) {
                    reject(result.error);
                } else {
                    resolve();
                }
            });
        });
    }

    public clear(): Promise<void> {
        return new Promise<void>((resolve, reject) => {
            this.remove(RoamingSettings02.CONFIGS_KEY);
            Office.context.roamingSettings.saveAsync((result: Office.AsyncResult) => {
                if (result.error) {
                    reject(result.error);
                } else {
                    RoamingSettings.ResetInstance();
                    resolve();
                }
            });
        });
    }

    private preload() {
        this.configs = [];
    }

    private populate() {
        this.configs = this.get<IVSTSConfig[]>(RoamingSettings02.CONFIGS_KEY);
    }

    private migrateFromRS01() {
        let rs: RoamingSettings01 = new RoamingSettings01();
        this.configs = [{
            account: rs.account,
            name: rs.team,
            project: rs.project,
            team: rs.team,
        }];
        rs.clear();
    }
}

export class RoamingSettings extends RoamingSettings01 {

    public static GetInstance(): RoamingSettings {
        if (!RoamingSettings.instance) {

            RoamingSettings.instance = new RoamingSettings01();
        }
        return RoamingSettings.instance;
    }

    public static ResetInstance(): void {
        RoamingSettings.instance = null;
    }

    // singleton instance
    private static instance: RoamingSettings;

}