/**
 * Brute force TypeScript type definition used by Office.js.
 */

declare module Office {
  interface AsyncResult {
    asynContex: any;
    error: any;
    status: any;
    value: any;
  }

  interface OfficeCallback { (asyncResult:AsyncResult): void }

  interface BodyInterface {
    getAsync(coersionType:string, options?:any, callback?: OfficeCallback);
  }

  interface ItemInterface {
    itemId: string;
    subject: string;
    normalizedSubject: string;
    body: BodyInterface;
    notificationMessages: any;
    displayReplyAllForm(form: any);
  }

  interface MailboxInterface {
    ewsUrl: string;
    getCallbackTokenAsync(callback: OfficeCallback, userContext?: any): void;
    item: ItemInterface;
    userProfile: any; 
  }

  export function initialize():any;

  export var cast:any;
  interface ContextInterface {
    mailbox: MailboxInterface;
    roamingSettings: any;
  }

  export var context: ContextInterface;

  export namespace MailboxEnums{
    export class ItemNotificationMessageType{
      static ProgressIndicator: string;
      static InformationalMessage: string;
      static ErrorMessage: string;
    }
  }
}

declare module 'office' {
  var out:typeof Office;
  export = out;
}

