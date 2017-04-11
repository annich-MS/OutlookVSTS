import * as React from "react";
import { CommandButton } from "office-ui-fabric-react/lib/Button";
import { Spinner } from "office-ui-fabric-react/lib/Spinner";
import { Rest } from "../../utils/rest";
import Constants from "../../models/constants";
import NavigationStore from "../../stores/navigationStore";
import { AppNotificationType } from "../../models/appNotification";

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
      <CommandButton icon="ReplyAll" onClick={this.handleClick.bind(this)}>Reply All with Work Item</CommandButton>);

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
  private async handleClick(): Promise<void> {
    if (Office.context.mailbox.diagnostics.hostName === Constants.IOS_HOST_NAME) {
      this.setState({ saving: true });
      try {
        await Rest.autoReply("I have created the following bug:<br/><br/>" + this.addSignature(this.props.workItemHyperlink));
        this.props.navigationStore.updateNotification({ message: "Message Sent!", type: AppNotificationType.Success });
      } catch (error) {
        this.props.navigationStore.updateNotification({ message: "Message Send Failed", type: AppNotificationType.Error });
      } finally {

        this.setState({ saving: false });
      }
    } else {
      (Office.context.mailbox.item as Office.MessageRead).displayReplyAllForm(this.addSignature(this.props.workItemHyperlink));
    }
  }
}
