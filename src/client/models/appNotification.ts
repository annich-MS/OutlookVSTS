export enum AppNotificationType {
    Error,
    Warning,
    Success,
    Assert,
}

export interface IAppNotification {
    type: AppNotificationType;
    message: string;
}

export default IAppNotification;
