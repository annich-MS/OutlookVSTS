export interface ServerTokenData {
    access_token: string;
    expires_in: string;
    refresh_token: string;
}

export default class Token {
    public static readonly TableName: string = "Tokens";
    public static readonly IdKey: string = "id";
    public static readonly TokenKey: string = "token";
    public static readonly ExpiryKey: string = "expiry";
    public static readonly RefreshKey: string = "refresh";

    public static getInstance(id: string, data: ServerTokenData): Token {
        let expiry: number = Date.now() + (Number(data.expires_in) * 1000); // data.expiresIn returns as seconds in string
        return new Token(id, data.access_token, expiry, data.refresh_token);
    }

    public readonly id: string;
    public readonly token: string;
    public readonly expiry: number;
    public readonly refresh: string;

    private constructor(id: string, token: string, expiry: number, refresh: string) {
        this.id = id;
        this.token = token;
        this.expiry = expiry;
        this.refresh = refresh;
    }

}
