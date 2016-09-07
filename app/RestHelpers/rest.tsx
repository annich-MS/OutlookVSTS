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

    public constructor(blob: any) {
        this.id = blob.accountId;
        this.name = blob.accountName;
        this.uri = blob.accountUri;
    }
}

export class Team {
    public id: string;
    public name: string;

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

interface IRestCallback { (output: string): void; }
interface IItemCallback { (item: string): void; }
interface IUserProfileCallback { (profile: UserProfile): void; }
interface IProjectsCallback { (projects: Project[]): void; }
interface IAccountsCallback { (accounts: Account[]): void; }
interface ITeamsCallback { (teams: Team[]): void; }
interface IWorkItemCallback { (workItemInfo: WorkItemInfo): void; }

export class Rest {

    private static userProfile: UserProfile;
    private static accounts: Account[];
    private static workItemInfo: WorkItemInfo;

    public static getItem(user: string, item: number, callback: IItemCallback): void {
        this.makeRestCallWithArgs('getItem', user, { fields: 'System.TeamProject', ids: item, instance: 'o365exchange' }, (output) => {
            callback(output);
        });
    }

    public static getUserProfile(user: string, callback: IUserProfileCallback): void {
        this.makeRestCall('me', user, (output) => {
            // console.log('get user prof' + output);
            this.userProfile = new UserProfile(JSON.parse(output));
            callback(this.userProfile);
        });
    }

    public static getAccounts(user: string, memberId: string, callback: IAccountsCallback): void {
        this.makeRestCallWithArgs('accounts', user, { memberId: memberId }, (output) => {
            let parsed: any = JSON.parse(output);
            this.accounts = [];
            parsed.value.forEach(account => {
                this.accounts.push(new Account(account));
            });
            callback(this.accounts);
        });
    }

    public static getProjects(user: string, accountName: string, callback: IProjectsCallback): void {
        this.makeRestCallWithArgs('projects', user, { account: accountName }, (output) => {
            let parsed: any = JSON.parse(output);
            let projects: Project[] = [];
            parsed.value.forEach(project => {
                projects.push(new Project(project));
            });
            callback(projects);
        });
    }

    public static getTeams(user: string, projectName: string, accountName: string, callback: ITeamsCallback): void {
        this.makeRestCallWithArgs('getTeams', user, { account: accountName, project: projectName }, (output) => {
            let parsed: any = JSON.parse(output);
            let teams: Team[] = [];
            parsed.value.forEach(team => {
                teams.push(new Team(team));
            });
            callback(teams);
        });
    }

    public static getTeamAreaPath(user: string, account: string, project: string, teamName: string, callback: IRestCallback): void {
        this.getTeams(user, project, account, (teams: Team[]) => {
            let guid: string;
            teams.forEach(team => {
                if (team.name === teamName) {
                    guid = team.id;
                }
            });
            this.makeRestCallWithArgs('getTeamField', user, { account: account, project: project, team: guid }, (output) => {
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

    public static getCurrentIteration(user: string, teamName: string, project: string, account: string, callback: IRestCallback): void {
        
        this.getTeams(user, project, account, (teams: Team[]) => {
            let guid: string;
            teams.forEach(team => {
                if (team.name === teamName) {
                    guid = team.id;
                }
            });
            this.makeRestCallWithArgs( 'getCurrentIteration', user, { account: account, project: project, team: guid }, (output) => {
                    callback(JSON.parse(output).value[0].path);
            });
        });
    }

    public static getMessage(user: string, ewsId: string, url: string, token: string, callback: IRestCallback ): void {
        Rest.makeRestCallWithArgs('getMessage', user, { ewsId:ewsId, url:url, token:token }, callback);
    }

    public static uploadAttachment(user: string, data: string, account: string, filename: string, callback: IRestCallback): void {
        Rest.makePostRestCallWithArgs('uploadAttachment', user, { account: account, filename: filename}, data, (output) =>
        { 
            var result = JSON.parse(output);
            callback(result.url);
        });
    }

    public static attachAttachment(user: string, account: any, attachmenturl: string, id: string, callback: IRestCallback): void {
        Rest.makeRestCallWithArgs('attachAttachment', user, { account: account, attachmenturl: attachmenturl, id: id }, callback);
    }

    public static createTask(user:string, options: any, account:string, project:string, team:string, callback: IWorkItemCallback): void {
        this.getTeamAreaPath(user, account, project, team, (areapath) => {
            options.areapath = areapath;
            options.account = account;
            options.project = project;
            options.team = team;
            this.getCurrentIteration(user, team, project, account, (iteration) => {
                options.iteration = iteration;
                this.makeRestCallWithArgs('createTask', user, options, (output) => {
                    callback(new WorkItemInfo(JSON.parse(output)));
                });
            });
        });
    }

    private static makeRestCall(name: string, user: string, callback: IRestCallback): void {
        $.get('./rest/' + name + '?user=' + user, callback);
    }

    private static makeRestCallWithArgs(name: string, user: string, args: any, callback: IRestCallback): void {
        const path: string = './rest/' + name + '?user=' + user + '&' + $.param(args);
        $.get(path, callback);
    }

    private static makePostRestCallWithArgs(name: string, user: string, args: any, body: string, callback: IRestCallback): void {
        let options = {
             url: "/rest/" + name + '?user=' + user + '&' + $.param(args),
             method: "POST",
             headers: {
                 "Content-Type": "text/plain"
             },
             data: body
        };
        $.ajax(options).done(callback);
    }

}
