import * as React from "react";
import { Button, ButtonType, Spinner } from "office-ui-fabric-react";
import { Rest } from "../rest";
import Constants from "../models/constants";
import NavigationStore from "../stores/navigationStore";
import { AppNotificationType } from "../models/appNotification";

/**
 * Props for ReplyAllButton Component
 */
interface IReplyAllButtonProps {
  /**
   * workItemHyperlink
   */
  workItemHyperlink: string;

  navigationStore: NavigationStore;
}

/**
 * Renders a button that on-click, opens a reply-all form with the item hyperlink inserted in-line
 */
export class ReplyAllButton extends React.Component<IReplyAllButtonProps, { saving: boolean }> {

  public constructor() {
    super();
    this.state = { saving: false };
  }

  /**
   * Renders the ReplyAllButton Component and reads IReplyAllButtonProps
   */
  public render(): JSX.Element {

    let item: any = (
      <Button
        buttonType={ButtonType.command}
        icon="ReplyAll"
        onClick={this.handleClick.bind(this)}>
        Reply All with Work Item
      </Button>);

    if (this.state.saving) {
      item = <Spinner />;
    }

    return (
      <div>
        {item}
      </div>
    );
  }

  /**
   * Adds signature line to the HTML body
   */
  public addSignature(workItemHyperlink: string): string {
    return workItemHyperlink + Constants.CREATED_STRING;
  }
  /**
   * Handles the click and displays a reply-all form
   */
  private handleClick(): void {
    if (Office.context.mailbox.diagnostics.hostName === Constants.IOS_HOST_NAME) {
      this.setState({ saving: true });
      Rest.autoReply("I have created the following bug:<br/><br/>" + this.addSignature(this.props.workItemHyperlink), (output: string) => {
        this.setState({ saving: false });
        this.props.navigationStore.updateNotification({ message: "Message Sent!", type: AppNotificationType.Success });
      });
    } else {
      (Office.context.mailbox.item as Office.MessageRead).displayReplyAllForm(this.addSignature(this.props.workItemHyperlink));
    }
  }
}
