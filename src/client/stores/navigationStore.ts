import { action, observable, computed } from "mobx";

import IAppNotification from "../models/appNotification";
import NavigationPage from "../models/navigationPage";

export default class NavigationStore {

    @computed get currentPage(): NavigationPage { return this._currentPage; };
    @computed get isSaving(): boolean { return this._isSaving; };
    @computed get notification(): IAppNotification { return this._notification; };

    @observable private _currentPage: NavigationPage = NavigationPage.Connecting;
    @observable private _lastPage: NavigationPage = NavigationPage.CreateWorkItem;
    @observable private _isSaving: boolean = false;
    @observable private _notification: IAppNotification = null;

    @action public navigate(newPage: NavigationPage, updateLast: boolean = true): void {
        if (updateLast) {
            this._lastPage = this._currentPage;
        }
        this._currentPage = newPage;
    }

    @action public navigateBack() {
        this._currentPage = this._lastPage;
    }

    @action public updateNotification(notification: IAppNotification): void {
        this._notification = notification;
    }

    @action public clearNotification(): void {
        this._notification = null;
    }

    @action public startSave(): void {
        this._isSaving = true;
    }

    @action public endSave(success: boolean) {
        this._isSaving = false;
        if (success) {
            this.navigate(NavigationPage.QuickActions);
        }
    }
}

export const navigationStore: NavigationStore = new NavigationStore();
