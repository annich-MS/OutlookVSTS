import * as React from "react";
import { observer } from "mobx-react";
import { computed } from "mobx";
import { Button, ButtonType } from "office-ui-fabric-react";

import { Rest } from "../rest";

import WorkItemStore from "../stores/workItemStore";
import NavigationStore from "../stores/navigationStore";
import APTCache from "../stores/aptCache";

import { AppNotificationType } from "../models/appNotification";
import APTPopulateStage from "../models/aptPopulateStage";
import { typeToString } from "../models/workItemType";
import VSTSInfo from "../models/vstsInfo";

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
  public async handleSave(): Promise<void> {
    try {
      let url: string = null;
      this.props.navigationStore.startSave();
      if (this.props.workItem.attachEmail) {
        url = await this.uploadAttachment();
      }
      await this.createWorkItem(url);
      this.props.navigationStore.endSave(true);
    } catch (e) {
      this.props.navigationStore.updateNotification(e);
      this.props.navigationStore.endSave(false);
    }
  }

  public async uploadAttachment(): Promise<string> {
    let id: string = (Office.context.mailbox.item as Office.MessageRead).itemId;
    let url: string = Office.context.mailbox.ewsUrl || "https://outlook.office365.com/EWS/Exchange.asmx";
    let token: string = await Rest.getCallbackToken();
    let message: string;
    try {
      message = await Rest.getMessage(id, url, token);
    } catch (e) { throw { message: e.toString("download message from Exchange"), type: AppNotificationType.Error }; }

    try {
      return await Rest.uploadAttachment(message, this.props.cache.account, `${(Office.context.mailbox.item as Message).normalizedSubject}.eml`);
    } catch (e) { throw { message: e.toString("upload attachment"), type: AppNotificationType.Error }; }
  }

  public async createWorkItem(attachmentUrl: string): Promise<void> {
    let options: any = {
      account: this.props.cache.account,
      attachment: attachmentUrl,
      project: this.props.cache.project,
      team: this.props.cache.team,
      title: this.props.workItem.title,
      type: typeToString(this.props.workItem.type),
    };
    let body: string = this.props.workItem.description;

    try {
      let info: VSTSInfo = await Rest.createTask(options, body);
      this.props.workItem.setInfo(info);
    } catch (error) {
      throw { message: error.toString("create task"), type: AppNotificationType.Error };
    }
    return;
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
