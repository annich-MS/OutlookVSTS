import { action, observable, computed } from "mobx";

import Constants from "../models/constants";
import WorkItemType from "../models/workItemType";
import VSTSInfo from "../models/vstsInfo";

export default class WorkItemStore {
    private static readonly MORE_DETAILS_STRING: string = `For more details, please refer to the attached mail thread.`;
    private static readonly DEFAULT_ATTACH_EMAIL: boolean = true;
    private static readonly DEFAULT_TYPE: WorkItemType = WorkItemType.Bug;

    @computed get title(): string { return this._title; };
    @computed get description(): string { return this._description; };
    @computed get attachEmail(): boolean { return this._attachEmail; };
    @computed get type(): WorkItemType { return this._type; };
    @computed get vstsInfo(): VSTSInfo { return this._vstsInfo; };

    @observable private _title: string = "";
    @observable private _description: string = `${WorkItemStore.MORE_DETAILS_STRING}${Constants.CREATED_STRING}`;
    @observable private _attachEmail: boolean = WorkItemStore.DEFAULT_ATTACH_EMAIL;
    @observable private _type: WorkItemType = WorkItemStore.DEFAULT_TYPE;
    @observable private _vstsInfo: VSTSInfo = null;

    @action public setTitle(title: string): void {
        this._title = title;
    }

    @action public setDescription(description: string): void {
        this._description = description;
    }

    @action public toggleAttachEmail(): void {
        this._attachEmail = !this._attachEmail;
        if (this._attachEmail) {
            this._description = `${WorkItemStore.MORE_DETAILS_STRING} ${this._description}`;
        } else {
            this._description = this._description.replace(WorkItemStore.MORE_DETAILS_STRING, ``);
        }
    }

    @action public setType(type: WorkItemType): void {
        this._type = type;
    }

    @action public setInfo(info: VSTSInfo): void {
        this._vstsInfo = info;
    }
}

export const workItemStore: WorkItemStore = new WorkItemStore();
