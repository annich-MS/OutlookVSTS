import * as Agent from "superagent";
import VSTSInfo from "./models/vstsInfo";

export class UserProfile {
    public displayName: string;
    public publicAlias: string;
    public emailAddress: string;
    public coreRevision: number;
    public timeStamp: string;
    public id: string;
    public revision: string;

    public constructor(blob: any) {
        this.displayName = blob.displayName;
        this.publicAlias = blob.publicAlias;
        this.id = blob.id;
    }
}

export abstract class DropdownParseable {
    public static compare(a: Project, b: Project): number {
        return a.name.toLowerCase() < b.name.toLowerCase() ? -1 : 1;
    }

    public name: string;
    public id: string;
}

export class Project extends DropdownParseable {
    public description: string;
    public url: string;
    public state: string;

    public constructor(blob: any) {
        super();
        this.id = blob.id;
        this.name = blob.name;
        this.description = blob.description;
        this.url = blob.url;
        this.state = blob.state;
    }

}

export class Account extends DropdownParseable {
    public uri: string;

    public constructor(blob: any) {
        super();
        this.id = blob.accountId;
        this.name = blob.accountName;
        this.uri = blob.accountUri;
    }
}

export class Team extends DropdownParseable {

    public constructor(blob: any) {
        super();
        this.id = blob.id;
        this.name = blob.name;
    }
}

interface VSTSErrorBody {
    message: string;
    typeKey: string;
}

export class RestError {

    public type: string;
    public more: any;
    public body: VSTSErrorBody;

    public constructor(blob: any) {
        this.type = blob.type;
        this.more = blob.more;
        try {
            this.body = JSON.parse(this.more.response.body);
        } catch (e) {
            // not parsable
        }
    }

    public toString(action: string): string {
        let reason: string = "";
        if (this.body) {
            reason = `${this.body.typeKey}: ${this.body.message}`;
        } else if (this.more.statusCode) {
            reason = `${this.more.name}. Server returned ${this.more.statusCode}`;
        } else {
            reason = `${this.type}.`;
        }
        return `Failed to ${action} due to ${reason}`;
    }
}

export abstract class Rest {

    public static async getIsAuthenticated(): Promise<boolean> {
        let user: string = await Rest.getUser();
        let res: Agent.Response = await Agent.get(`./authenticate/db`)
            // zzGarbage property to prevent IE caching the result
            .query({ user: user, zzGarbage: Math.random() * 1000 });
        return res.text === "success";
    }

    public static async getItem(item: number): Promise<string> {
        return await this.makeRestCallWithArgs("getItem", {
            fields: "System.TeamProject",
            ids: item,
            instance: "o365exchange",
        });
    }


    public static async getUserProfile(): Promise<UserProfile> {
        let output: string = await this.makeRestCall("me");

        let parsed: any = JSON.parse(output);
        if (parsed.error) {
            throw new RestError(parsed.error);
        }
        this.userProfile = new UserProfile(parsed);
        return this.userProfile;
    }

    public static async getAccounts(memberId: string): Promise<Account[]> {
        let output: string = await this.makeRestCallWithArgs("accounts", { memberId: memberId });
        let parsed: { error?: any, value: any[] } = JSON.parse(output);
        if (parsed.error) {
            throw new RestError(parsed.error);
        }
        return parsed.value.map(account => { return new Account(account); });
    }

    public static async getProjects(accountName: string): Promise<Project[]> {
        let output: string = await this.makeRestCallWithArgs("projects", { account: accountName });
        let parsed: { error?: any, value: any[] } = JSON.parse(output);
        if (parsed.error) {
            throw new RestError(parsed.error);
        }
        return parsed.value.map(project => { return new Project(project); });
    }

    public static async getTeams(projectName: string, accountName: string): Promise<Team[]> {
        let output: string = await this.makeRestCallWithArgs("getTeams", { account: accountName, project: projectName });
        let parsed: { error?: any, value: any[] } = JSON.parse(output);
        if (parsed.error) {
            throw new RestError(parsed.error);
        }
        return parsed.value.map(team => { return new Team(team); });
    }

    public static async getIteration(teamName: string, project: string, account: string): Promise<string> {
        let teams: Team[] = await this.getTeams(project, account);
        let guid: string = teams.filter(team => { return team.name === teamName; })[0].id;
        let output: string = await this.makeRestCallWithArgs("backlog", { account: account, project: project, team: guid });
        console.log(output);
        let parsed: any = JSON.parse(output);
        if (parsed.error) {
            throw new RestError(parsed.error);
        }
        if (parsed.backlogIteration.id !== "00000000-0000-0000-0000-000000000000") {
            return parsed.backlogIteration.path;
        }
        throw new RestError({ more: "Missing Backlog Iteration", type: "Missing Backlog Iteration" });
    }

    public static async getTeamAreaPath(account: string, project: string, teamName: string): Promise<string> {
        let output: string = await this.makeRestCallWithArgs("getTeamField", { account: account, project: project, team: teamName });
        let parsed: any = JSON.parse(output);
        if (parsed.error) {
            throw new RestError(parsed.error);
        }
        if (parsed.field.referenceName !== "System.AreaPath") {
            // we don"t support teams that don't use area path as their team field
            throw new RestError({
                more: "The vsts add-in does not support teams that do not have area path as their default field",
                type: "Missing Area Path",
            });
        }
        return parsed.defaultValue;
    }


    public static async getMessage(ewsId: string, url: string, token: string): Promise<string> {
        let output: string = await Rest.makeRestCallWithArgs("getMessage", { ewsId: ewsId, token: token, url: url });
        let parsed: { error?: any } = null;
        try {
            parsed = JSON.parse(output); // will only succeed if error returned
        } catch (e) {
            return output;
        }

        throw new RestError(parsed.error);
    }

    public static async uploadAttachment(data: string, account: string, filename: string): Promise<string> {
        let output: string = await Rest.makePostRestCallWithArgs("uploadAttachment", { account: account, filename: filename }, data);
        let parsed: { error?: any, url: string } = JSON.parse(output);
        if (parsed.error) {
            throw new RestError(parsed.error);
        }
        return parsed.url;
    }

    public static async createTask(options: any, body: string): Promise<VSTSInfo> {
        options.areapath = await this.getTeamAreaPath(options.account, options.project, options.team);
        options.iteration = await this.getIteration(options.team, options.project, options.account);
        let output: string = await this.makePostRestCallWithArgs("createTask", options, body);
        let parsed: any = JSON.parse(output);
        if (parsed.error) {
            throw new RestError(parsed.error);
        }
        return new VSTSInfo(parsed);
    }

    public static async removeUser(): Promise<void> {
        let output: string = await Rest.makeRestCall("disconnect");
        let parsed: any = JSON.parse(output);
        if (parsed.error) {
            throw new RestError(parsed.error);
        } else {
            return;
        }
    }

    public static getUser(): Promise<string> {
        return new Promise<string>((resolve, reject) => {
            Rest.tokenRequests++;
            if (Rest.uidToken !== "" && Rest.tokenRefresh.getTime() > Date.now()) {
                resolve(Rest.uidToken);
                Rest.tokenCacheHits++;
            } else {
                Office.context.mailbox.getUserIdentityTokenAsync((asyncResult: Office.AsyncResult) => {
                    if (asyncResult.error) {
                        reject(asyncResult.error);
                    }
                    Rest.uidToken = asyncResult.value;
                    Rest.tokenRefresh = new Date(Date.now() + 10 * 60 * 1000); // 10 * sec/min * ms/sec
                    resolve(asyncResult.value);
                });
            }
        });
    }

    public static getCallbackToken(): Promise<string> {
        return new Promise<string>((resolve, reject) => {
            Office.context.mailbox.getCallbackTokenAsync((asyncResult: Office.AsyncResult) => {
                if (asyncResult.error) {
                    reject(asyncResult.error);
                } else {
                    resolve(asyncResult.value);
                }
            });
        });
    }

    public static async autoReply(msg: string): Promise<void> {
        let token: string = await Rest.getCallbackToken();
        let args: any = {
            item: (Office.context.mailbox.item as Office.ItemRead).itemId,
            token: token,
        };
        let body: string = JSON.stringify({ "Comment": msg });
        await Rest.makePostRestCallWithArgs("reply", args, body);
    }

    public static log(msg: string): void {
        Agent.get(`./log?msg=${encodeURIComponent(msg)}`).end();
    }

    private static userProfile: UserProfile;
    private static uidToken: string = "";
    private static tokenRefresh: Date = new Date();
    private static tokenRequests: number = 0;
    private static tokenCacheHits: number = 0;

    private static makeRestCall(name: string): Promise<string> {
        return Rest.makeRestCallWithArgs(name, {});
    }

    private static async makeRestCallWithArgs(name: string, args: any): Promise<string> {
        let user: string = await Rest.getUser();
        args.user = user;
        // the randomized element should prevent IE from caching the response.
        args.ieRandomizer = Math.floor(Math.random() * 100000);

        let res: Agent.Response = await Agent.get(`./rest/${name}`).query(args);
        return res.text;
    }

    private static async makePostRestCallWithArgs(name: string, args: any, body: string): Promise<string> {
        args.user = await Rest.getUser();
        let res: Agent.Response = await Agent.post(`./rest/${name}`)
            .query(args)
            .set("Content-Type", "text/plain")
            .send(body);
        return res.text;
    }

}
