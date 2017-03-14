export default class AuthInfo {

    public static getInstance(): AuthInfo {
        if (AuthInfo.instance === undefined) {
            AuthInfo.instance = new AuthInfo();
        }
        return AuthInfo.instance;
    }

    private static instance: AuthInfo;

    // constants
    public readonly authEndpoint: string = `oauth2/authorize`;
    public readonly tokenEndpoint: string = `oauth2/token`;
    public readonly baseUrl: string = `app.vssps.visualstudio.com`;

    public readonly id: string;
    public readonly secret: string;
    public readonly scopes: string;
    public readonly redirect: string;

    private constructor() {
        this.id = process.env.VSTS_ID;
        this.secret = process.env.VSTS_SECRET;
        this.scopes = process.env.VSTS_SCOPES;
        this.redirect = process.env.VSTS_REDIRECT;
    }
}
