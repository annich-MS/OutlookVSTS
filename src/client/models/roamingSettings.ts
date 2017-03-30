import { IDropdownOption } from "office-ui-fabric-react";
import IVSTSConfig from "../models/vstsConfig";
import APTCache from "../stores/aptCache";

abstract class BaseRoamingSettings {
    protected static readonly VERSION_KEY: string = "version";

    protected static readonly ID_KEY: string = "member_id";

    public id: string;

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

    public static readonly VERSION: number = 1;

    public static async _getInstance(): Promise<RoamingSettings01> {
        let rs: RoamingSettings01 = new RoamingSettings01();
        let version: number = rs.get<number>(RoamingSettings.VERSION_KEY);
        if (version === undefined || version < RoamingSettings01.VERSION) {
            rs.isValid = false;
        } else {
            rs.isValid = true;
            rs.account = rs.get<string>(RoamingSettings01.ACCOUNT_KEY) || "";
            rs.project = rs.get<string>(RoamingSettings01.PROJECT_KEY) || "";
            rs.team = rs.get<string>(RoamingSettings01.TEAM_KEY) || "";
            rs.accounts = rs.get<IDropdownOption[]>(RoamingSettings01.ACCOUNTS_KEY) || [];
            rs.projects = rs.get<IDropdownOption[]>(RoamingSettings01.PROJECTS_KEY) || [];
            rs.teams = rs.get<IDropdownOption[]>(RoamingSettings01.TEAMS_KEY) || [];
            rs.id = rs.get<string>(RoamingSettings.ID_KEY);
        }
        return rs;
    }

    // constants
    private static readonly ACCOUNT_KEY: string = "default_account";
    private static readonly PROJECT_KEY: string = "default_project";
    private static readonly TEAM_KEY: string = "default_team";
    private static readonly ACCOUNTS_KEY: string = "accounts";
    private static readonly PROJECTS_KEY: string = "projects";
    private static readonly TEAMS_KEY: string = "teams";

    public isValid: boolean = false;
    public isFull: boolean = false;
    public account: string = "";
    public project: string = "";
    public team: string = "";
    public accounts: IDropdownOption[] = [];
    public projects: IDropdownOption[] = [];
    public teams: IDropdownOption[] = [];

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
            this.set(RoamingSettings.ID_KEY, this.id);
            this.set(RoamingSettings01.ACCOUNT_KEY, this.account);
            this.set(RoamingSettings01.PROJECT_KEY, this.project);
            this.set(RoamingSettings01.TEAM_KEY, this.team);
            this.set(RoamingSettings01.ACCOUNTS_KEY, this.accounts);
            this.set(RoamingSettings01.PROJECTS_KEY, this.projects);
            this.set(RoamingSettings01.TEAMS_KEY, this.teams);
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
            this.remove(RoamingSettings.ID_KEY);
            this.remove(RoamingSettings01.ACCOUNT_KEY);
            this.remove(RoamingSettings01.PROJECT_KEY);
            this.remove(RoamingSettings01.TEAM_KEY);
            this.remove(RoamingSettings01.ACCOUNTS_KEY);
            this.remove(RoamingSettings01.PROJECTS_KEY);
            this.remove(RoamingSettings01.TEAMS_KEY);
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

    public static async _getInstance(): Promise<RoamingSettings02> {
        let rs: RoamingSettings02 = new RoamingSettings02();
        let version: number = rs.get<number>(RoamingSettings.VERSION_KEY);
        switch (version) {
            case RoamingSettings01.VERSION:
                await rs.migrateFromRS01();
            case RoamingSettings02.VERSION:
                rs.fromRoamingSettings();
            default:
                rs.preload();
        }
        return rs;
    }

    private static readonly CONFIGS_KEY: string = "configs";

    public configs: IVSTSConfig[];

    public save(): Promise<void> {
        return new Promise<void>((resolve, reject) => {
            console.log("saving " + JSON.stringify(this));
            try {
                this.set<IVSTSConfig[]>(RoamingSettings02.CONFIGS_KEY, this.configs);
                this.set<number>(RoamingSettings.VERSION_KEY, RoamingSettings02.VERSION);
                this.set<string>(RoamingSettings.ID_KEY, this.id);
                Office.context.roamingSettings.saveAsync((result: Office.AsyncResult) => {
                    console.log("Hello");
                    if (result.error) {
                        console.log("Error " + JSON.stringify(result.error));
                        reject(result.error);
                    } else {
                        console.log("Success");
                        resolve();
                    }
                });
            } catch (e) {
                console.log(e);
            }
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

    private preload(): void {
        this.configs = [];
    }

    private fromRoamingSettings(): void {
        console.log("At get:" + this.get<IVSTSConfig[]>(RoamingSettings02.CONFIGS_KEY));
        this.configs = this.get<IVSTSConfig[]>(RoamingSettings02.CONFIGS_KEY);
        this.id = this.get<string>(RoamingSettings.ID_KEY);
    }

    private async migrateFromRS01(): Promise<void> {
        let rs: RoamingSettings01 = await RoamingSettings01._getInstance();
        this.configs = [{
            account: rs.account,
            name: "Default",
            project: rs.project,
            team: rs.team,
        }];
        await rs.clear();
        return;
    }
}

export default class RoamingSettings extends RoamingSettings02 {

    public static async GetInstance(): Promise<RoamingSettings> {
        if (!RoamingSettings.instance) {

            RoamingSettings.instance = (await RoamingSettings02._getInstance()) as RoamingSettings;
        }
        return RoamingSettings.instance;
    }

    public static ResetInstance(): void {
        RoamingSettings.instance = null;
    }

    // singleton instance
    private static instance: RoamingSettings;

}
