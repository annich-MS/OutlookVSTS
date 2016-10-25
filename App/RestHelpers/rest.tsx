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

export class Project {
    public id: string;
    public name: string;
    public description: string;
    public url: string;
    public state: string;

    public static compare(a: Project, b: Project): number {
        return a.name.toLowerCase() < b.name.toLowerCase() ? -1 : 1;
    }

    public constructor(blob: any) {
        this.id = blob.id;
        this.name = blob.name;
        this.description = blob.description;
        this.url = blob.url;
        this.state = blob.state;
    }

}

export class Account {
    public id: string;
    public name: string;
    public uri: string;

    public static compare(a: Account, b: Account): number {
        return a.name.toLowerCase() < b.name.toLowerCase() ? -1 : 1;
    }

    public constructor(blob: any) {
        this.id = blob.accountId;
        this.name = blob.accountName;
        this.uri = blob.accountUri;
    }
}

export class Team {
    public id: string;
    public name: string;

    public static compare(a: Team, b: Team): number {
        return a.name.toLowerCase() < b.name.toLowerCase() ? -1 : 1;
    }

    public constructor(blob: any) {
        this.id = blob.id;
        this.name = blob.name;
    }
}

export class WorkItemInfo {
    public id: string;
    public VSTShtmlLink: string;

    public constructor(blob: any) {
        this.id = blob.id;
        this.VSTShtmlLink = blob._links.html.href;
    }
}

export class RestError {
    public type: string;
    public more: string;

    public constructor(blob: any) {
        this.type = blob.type;
        this.more = blob.more;
    }
}

export interface IRestCallback { (output: string): void; }
interface IItemCallback { (item: string): void; }
interface IErrorCallback { (error: RestError): void; }
interface IUserProfileCallback { (profile: UserProfile): void; }
interface IProjectsCallback { (projects: Project[]): void; }
interface IAccountsCallback { (accounts: Account[]): void; }
interface ITeamsCallback { (teams: Team[]): void; }
interface IWorkItemCallback { (workItemInfo: WorkItemInfo): void; }

export abstract class Rest {

    private static userProfile: UserProfile;
    private static accounts: Account[];


    public static getItem(item: number, callback: IItemCallback): void {
        this.makeRestCallWithArgs('getItem', { fields: 'System.TeamProject', ids: item, instance: 'o365exchange' }, (output) => {
            callback(output);
        });
    }

    public static getUserProfile(callback: IUserProfileCallback): void {
        this.makeRestCall('me', (output) => {
            // console.log('get user prof' + output);
            this.userProfile = new UserProfile(JSON.parse(output));
            callback(this.userProfile);
        });
    }

    public static getAccounts(memberId: string, callback: IAccountsCallback): void {
        this.makeRestCallWithArgs('accounts', { memberId: memberId }, (output) => {
            let parsed: any = JSON.parse(output);
            this.accounts = [];
            parsed.value.forEach(account => {
                this.accounts.push(new Account(account));
            });
            callback(this.accounts);
        });
    }

    public static getProjects(accountName: string, callback: IProjectsCallback): void {
        this.makeRestCallWithArgs('projects', { account: accountName }, (output) => {
            let parsed: any = JSON.parse(output);
            let projects: Project[] = [];
            parsed.value.forEach(project => {
                projects.push(new Project(project));
            });
            callback(projects);
        });
    }

    public static getTeams(projectName: string, accountName: string, callback: ITeamsCallback): void {
        this.makeRestCallWithArgs('getTeams', { account: accountName, project: projectName }, (output) => {
            let parsed: any = JSON.parse(output);
            let teams: Team[] = [];
            parsed.value.forEach(team => {
                teams.push(new Team(team));
            });
            callback(teams);
        });
    }

    public static getTeamAreaPath(account: string, project: string, teamName: string, callback: IRestCallback): void {
        this.getTeams(project, account, (teams: Team[]) => {
            let guid: string;
            teams.forEach(team => {
                if (team.name === teamName) {
                    guid = team.id;
                }
            });
            this.makeRestCallWithArgs('getTeamField', { account: account, project: project, team: guid }, (output) => {
                let parsed: any = JSON.parse(output);
                if (parsed.field.referenceName !== 'System.AreaPath') {
                    // we don't support teams that don't use area path as their team field
                    callback('');
                } else {
                    callback(parsed.defaultValue);
                }
            });
        });
    }

    public static getCurrentIteration(teamName: string, project: string, account: string, callback: IRestCallback): void {

        this.getTeams(project, account, (teams: Team[]) => {
            let guid: string;
            teams.forEach(team => {
                if (team.name === teamName) {
                    guid = team.id;
                }
            });
            this.makeRestCallWithArgs('getCurrentIteration', { account: account, project: project, team: guid }, (output) => {
                callback(JSON.parse(output).value[0].path);
            });
        });
    }

    public static getMessage(ewsId: string, url: string, token: string, callback: IRestCallback): void {
        Rest.makeRestCallWithArgs('getMessage', { ewsId: ewsId, token: token, url: url }, callback);
    }

    public static uploadAttachment(data: string, account: string, filename: string, callback: IRestCallback): void {
        Rest.makePostRestCallWithArgs('uploadAttachment', { account: account, filename: filename }, data, (output) => {
            let result: any = JSON.parse(output);
            callback(result.url);
        });
    }

    public static attachAttachment(account: any, attachmenturl: string, id: string, callback: IRestCallback): void {
        Rest.makeRestCallWithArgs('attachAttachment', { account: account, attachmenturl: attachmenturl, id: id }, callback);
    }

    public static createTask(options: any, account: string, project: string, team: string, callback: IWorkItemCallback): void {
        this.getTeamAreaPath(account, project, team, (areapath) => {
            options.areapath = areapath;
            options.account = account;
            options.project = project;
            options.team = team;
            this.getCurrentIteration(team, project, account, (iteration) => {
                options.iteration = iteration;
                this.makeRestCallWithArgs('createTask', options, (output) => {
                    callback(new WorkItemInfo(JSON.parse(output)));
                });
            });
        });
    }

    public static removeUser(callback: IErrorCallback): void {
        Rest.makeRestCall('disconnect', (output) => {
            let parsed: any = JSON.parse(output);
            if (parsed.error) {
                callback(new RestError(parsed.error));
            } else {
                callback(null);
            }
        });
    }
    public static getUser(callback: IRestCallback): void {
        Office.context.mailbox.getUserIdentityTokenAsync((asyncResult: Office.AsyncResult) => {
            callback(asyncResult.value);
        });
    }

    private static makeRestCall(name: string, callback: IRestCallback): void {
        Rest.getUser((user: string) => {
            $.get('./rest/' + name + '?user=' + user, callback);
        });
    }

    private static makeRestCallWithArgs(name: string, args: any, callback: IRestCallback): void {
        Rest.getUser((user: string) => {
            const path: string = './rest/' + name + '?user=' + user + '&' + $.param(args);
            $.get(path, callback);
        });
    }

    private static makePostRestCallWithArgs(name: string, args: any, body: string, callback: IRestCallback): void {
        Rest.getUser((user: string) => {
            let options: any = {
                data: body,
                headers: {
                    'Content-Type': 'text/plain',
                },
                method: 'POST',
                url: '/rest/' + name + '?user=' + user + '&' + $.param(args),
            };
            $.ajax(options).done(callback);
        });
    }

}
