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
    public more: any;

    public constructor(blob: any) {
        this.type = blob.type;
        this.more = blob.more;
    }

    public toString(source: string): string {
        let contents: string = 'Failed to ' + source + ' due to ';
        if (this.more.statusCode) {
            contents += this.more.name + '. External server returned ' + this.more.statusCode;
        } else {
            contents += this.type + '.';
        }
        return contents;
    }
}

export interface IRestCallback { (output: string): void; }
interface IItemCallback { (error: RestError, item: string): void; }
export interface IStringCallback { (error: RestError, data: string): void; }
interface IErrorCallback { (error: RestError): void; }
interface IUserProfileCallback { (error: RestError, profile: UserProfile): void; }
interface IProjectsCallback { (error: RestError, projects: Project[]): void; }
interface IAccountsCallback { (error: RestError, accounts: Account[]): void; }
interface ITeamsCallback { (error: RestError, teams: Team[]): void; }
interface IWorkItemCallback { (error: RestError, workItemInfo: WorkItemInfo): void; }

export abstract class Rest {

    private static userProfile: UserProfile;
    private static accounts: Account[];
    private static uidToken: string = '';
    private static tokenRefresh: Date = new Date();
    private static tokenRequests: number = 0;
    private static tokenCacheHits: number = 0;


    public static getItem(item: number, callback: IItemCallback): void {
        this.makeRestCallWithArgs('getItem', { fields: 'System.TeamProject', ids: item, instance: 'o365exchange' }, (output) => {
            try {
                let parsed: any = JSON.parse(output); // will only succeed if error returned
                callback(new RestError(parsed.error), null);
            } catch (e) {
                callback(null, output);
            }
        });
    }


    public static getUserProfile(callback: IUserProfileCallback): void {
        this.makeRestCall('me', (output) => {
            let parsed: any = JSON.parse(output);
            if (parsed.error) {
                callback(new RestError(parsed.error), null);
                return;
            }
            this.userProfile = new UserProfile(parsed);
            callback(null, this.userProfile);
        });
    }

    public static getAccounts(memberId: string, callback: IAccountsCallback): void {
        this.makeRestCallWithArgs('accounts', { memberId: memberId }, (output) => {
            let parsed: any = JSON.parse(output);
            if (parsed.error) {
                callback(new RestError(parsed.error), null);
                return;
            }
            this.accounts = [];
            parsed.value.forEach(account => {
                this.accounts.push(new Account(account));
            });
            callback(null, this.accounts);
        });
    }

    public static getProjects(accountName: string, callback: IProjectsCallback): void {
        this.makeRestCallWithArgs('projects', { account: accountName }, (output) => {
            let parsed: any = JSON.parse(output);
            if (parsed.error) {
                callback(new RestError(parsed.error), null);
                return;
            }
            let projects: Project[] = [];
            parsed.value.forEach(project => {
                projects.push(new Project(project));
            });
            callback(null, projects);
        });
    }

    public static getTeams(projectName: string, accountName: string, callback: ITeamsCallback): void {
        this.makeRestCallWithArgs('getTeams', { account: accountName, project: projectName }, (output) => {
            let parsed: any = JSON.parse(output);
            if (parsed.error) {
                callback(new RestError(parsed.error), null);
                return;
            }
            let teams: Team[] = [];
            parsed.value.forEach(team => {
                teams.push(new Team(team));
            });
            callback(null, teams);
        });
    }


    public static getTeamAreaPath(account: string, project: string, teamName: string, callback: IStringCallback): void {
        this.makeRestCallWithArgs('getTeamField', { account: account, project: project, team: teamName }, (output) => {
            let parsed: any = JSON.parse(output);
            if (parsed.error) {
                callback(new RestError(parsed.error), null);
                return;
            }
            if (parsed.field.referenceName !== 'System.AreaPath') {
                // we don't support teams that don't use area path as their team field
                callback(null, '');
            } else {
                callback(null, parsed.defaultValue);
            }
        });
    }

    public static getIteration(teamName: string, project: string, account: string, callback: IStringCallback): void {
        this.getTeams(project, account, (error: RestError, teams: Team[]) => {
            if (error) {
                callback(error, null);
                return;
            }
            let guid: string;
            teams.forEach(team => {
                if (team.name === teamName) {
                    guid = team.id
                }
            });
            this.makeRestCallWithArgs('backlog', { account: account, project: project, team: guid }, (output) => {
                let parsed: any = JSON.parse(output);
                if (parsed.error) {
                    callback(new RestError(parsed.error), null);
                    return;
                }
                if (parsed.backlogIteration.id !== '00000000-0000-0000-0000-000000000000') {
                    callback(null, parsed.backlogIteration.path);
                } else {
                    callback(new RestError({ more: 'Missing Backlog Iteration', type: 'Missing Backlog Iteration'}), '');
                }
            });
        });
    }

    public static getMessage(ewsId: string, url: string, token: string, callback: IStringCallback): void {
        Rest.makeRestCallWithArgs('getMessage', { ewsId: ewsId, token: token, url: url }, (output) => {
            try {
                let parsed: any = JSON.parse(output); // will only succeed if error returned
                callback(new RestError(parsed.error), null);
            } catch (e) {
                callback(null, output);
            }
        });
    }

    public static uploadAttachment(data: string, account: string, filename: string, callback: IStringCallback): void {
        Rest.makePostRestCallWithArgs('uploadAttachment', { account: account, filename: filename }, data, (output) => {
            let parsed: any = JSON.parse(output);
            if (parsed.error) {
                callback(new RestError(parsed.error), null);
                return;
            }
            callback(null, parsed.url);
        });
    }

    public static attachAttachment(account: any, attachmenturl: string, id: string, callback: IStringCallback): void {
        Rest.makeRestCallWithArgs('attachAttachment', { account: account, attachmenturl: attachmenturl, id: id }, (output) => {
            try {
                let parsed: any = JSON.parse(output); // will only succeed if error returned
                callback(new RestError(parsed.error), null);
            } catch (e) {
                callback(null, output);
            }
        });
    }

    public static createTask(options: any, body: string, callback: IWorkItemCallback): void {
        this.getTeamAreaPath(options.account, options.project, options.team, (err, areapath) => {
            if (err) {
                callback(err, null);
                return;
            }
            options.areapath = areapath;
            this.getIteration(options.team, options.project, options.account, (err2, iteration) => {
                if (err2) {
                    callback(err2, null);
                }
                options.iteration = iteration;
                this.makePostRestCallWithArgs('createTask', options, body, (output) => {
                    let parsed: any = JSON.parse(output);
                    if (parsed.error) {
                        callback(new RestError(parsed.error), null);
                        return;
                    }
                    callback(null, new WorkItemInfo(parsed));
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
        Rest.tokenRequests++;
        if (Rest.uidToken !== '' && Rest.tokenRefresh.getTime() > Date.now()) {
            callback(Rest.uidToken);
            Rest.tokenCacheHits++;
        } else {
            Office.context.mailbox.getUserIdentityTokenAsync((asyncResult: Office.AsyncResult) => {
                Rest.uidToken = asyncResult.value;
                Rest.tokenRefresh = new Date(Date.now() + 10 * 60 * 1000); // 10 * sec/min * ms/sec
                callback(asyncResult.value);
            });
        }

    }

    public static autoReply(msg: string, callback: IRestCallback): void {
        Office.context.mailbox.getCallbackTokenAsync((asyncResult: Office.AsyncResult) => {
            let args: any = {
                item: Office.context.mailbox.item.itemId,
                token: asyncResult.value,
            };
            let body: string = JSON.stringify({'Comment': msg});
            Rest.makePostRestCallWithArgs('reply', args, body, callback);
        });
    }

    public static log(msg: string): void {
        $.get('./log?msg=' + encodeURIComponent(msg));
    }

    private static makeRestCall(name: string, callback: IRestCallback): void {
        Rest.makeRestCallWithArgs(name, {}, callback);
    }

    private static makeRestCallWithArgs(name: string, args: any, callback: IRestCallback): void {
        // the randomized element should prevent IE from caching the response.
        args.ieRandomizer = Math.floor(Math.random() * 100000);
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
