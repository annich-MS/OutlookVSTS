export default class Token {
    public static readonly TableName: string = "Tokens";
    public static readonly IdKey: string = "id";
    public static readonly TokenKey: string = "token";
    public static readonly ExpiryKey: string = "expiry";
    public static readonly RefreshKey: string = "refresh";

    public static getInstance(id: string, token: string, expiry: number, refresh: string) {
        return new Token(id, token, expiry, refresh);
    }

    public readonly id: string;
    public readonly token: string;
    public readonly expiry: Date;
    public readonly refresh: string;

    private constructor(id: string, token: string, expiry: number, refresh: string) {
        this.id = id;
        this.token = token;
        this.expiry = new Date(Date.now() + expiry);
        this.refresh = refresh;
    }

}
