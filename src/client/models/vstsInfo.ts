export default class VSTSInfo {
    public vstsUrl: string;
    public id: string;

    public constructor(blob: any) {
        this.id = blob.id;
        this.vstsUrl = blob._links.html.href;
    }
}