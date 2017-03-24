import * as React from "react";
import { Rest, RestError, WorkItemInfo, IStringCallback } from "../rest";
import { Button, ButtonType } from "office-ui-fabric-react";
import WorkItemStore from "../stores/workItemStore";
import NavigationStore from "../stores/navigationStore";
import APTCache from "../stores/aptCache";
import { AppNotificationType } from "../models/appNotification";
import APTPopulateStage from "../models/aptPopulateStage";

import { observer } from "mobx-react";
import { computed } from "mobx";

type Message = Office.MessageRead;

/**
 * Represents the Save Properties
 */
export interface ISaveProps {
  cache: APTCache;
  navigationStore: NavigationStore;
  workItem: WorkItemStore;
}

@observer
export class Save extends React.Component<ISaveProps, {}> {
  /**
   * Dispatches the action to change the Stage and make the REST call to create the work item
   * @returns {void}
   */
  public handleSave(): void {
    this.props.navigationStore.startSave();
    if (this.props.workItem.attachEmail) {
      Office.context.mailbox.getCallbackTokenAsync((tokenResult) => {
        this.uploadAttachment(tokenResult.value, (error, attachmentUrl) => { this.createWorkItem(attachmentUrl); });
      });
    } else {
      this.createWorkItem(null);
    }
  }

  public uploadAttachment(token: string, callback: IStringCallback): void {
    let id: string = (Office.context.mailbox.item as Office.MessageRead).itemId;
    let url: string = Office.context.mailbox.ewsUrl || "https://outlook.office365.com/EWS/Exchange.asmx";
    let account: string = this.props.cache.account;

    Rest.getMessage(id, url, token, (error, data) => {
      if (error) {
        this.props.navigationStore.updateNotification({ message: error.toString("download message from Exchange"), type: AppNotificationType.Error });
        this.props.navigationStore.endSave(false);
        return;
      }
      Rest.uploadAttachment(data, account, (Office.context.mailbox.item as Message).normalizedSubject + ".eml", (err, link) => {
        if (err) {
          this.props.navigationStore.updateNotification({ message: error.toString("upload attachment"), type: AppNotificationType.Error });
          this.props.navigationStore.endSave(false);
          return;
        }
        callback(null, link);
      });
    });

  }

  public createWorkItem(attachmentUrl: string): void {
    let options: any = {
      account: this.props.cache.account,
      attachment: attachmentUrl,
      project: this.props.cache.project,
      team: this.props.cache.team,
      title: this.props.workItem.title,
      type: this.props.workItem.type,
    };
    let body: string = this.props.workItem.description;


    Rest.createTask(options, body, (error: RestError, workItemInfo: WorkItemInfo) => {
      if (error) {
        this.props.navigationStore.updateNotification({ message: error.toString("create task"), type: AppNotificationType.Error });
        this.props.navigationStore.endSave(false);
        return;
      }
      this.props.navigationStore.endSave(true);
    });
  }


  /**
   * Renders the Save button and disables it on click
   */
  public render(): JSX.Element {

    let text: string = this.props.navigationStore.isSaving ? "Creating..." : "Create work item";
    return (
      <div style={{ textAlign: "center" }} >
        <br />
        <Button
          buttonType={ButtonType.primary}
          disabled={!this.shouldBeEnabled}
          onClick={this.handleSave.bind(this)} > {text} </Button>
      </div>
    );
  }


  @computed private get shouldBeEnabled(): boolean {
    return !(
      this.props.navigationStore.isSaving ||
      this.props.cache.populateStage < APTPopulateStage.PostPopulate ||
      this.props.navigationStore.notification != null);
  }
}
