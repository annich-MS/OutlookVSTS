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
  
  interface DiagnosticsInterface {
    hostName: string;
  }

  interface MailboxInterface {
    ewsUrl: string;
    displayNewMessageForm(messageData: Object): void;
    getCallbackTokenAsync(callback: OfficeCallback, userContext?: any): void;
    getUserIdentityTokenAsync(callback: OfficeCallback, userContext?: any): void;
    item: ItemInterface;
    userProfile: any; 
    diagnostics: DiagnosticsInterface;
  }

  export function initialize():any;

  export var cast:any;
  
  interface UiInterface {
    displayDialogAsync(url: string, options?: any, callback?: OfficeCallback);
  }

  interface ContextInterface {
    mailbox: MailboxInterface;
    roamingSettings: any;
    ui: UiInterface;
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

