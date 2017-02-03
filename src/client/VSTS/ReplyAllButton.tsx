import * as React from 'react';
import { Provider } from 'react-redux';
import { Button, ButtonType, Spinner, SpinnerType } from 'office-ui-fabric-react';
import { updateNotificationAction, NotificationType } from '../Redux/FlowActions';
import { Rest } from '../RestHelpers/rest';

/**
 * Props for ReplyAllButton Component
 * @interface { IReplyAllButtonProps }
 */
interface IReplyAllButtonProps {
  /**
   * workItemHyperlink
   * @type { string }
   */
  workItemHyperlink: string;

  dispatch?: Function;
}

/**
 * Renders a button that on-click, opens a reply-all form with the item hyperlink inserted in-line
 * @class { ReplyAllButton }
 */
export class ReplyAllButton extends React.Component<IReplyAllButtonProps, { saving: boolean }> {

  public constructor() {
    super();
    this.state = { saving: false };
  }

  /**
   * Renders the ReplyAllButton Component and reads IReplyAllButtonProps
   * @returns { React.ReactElement } ReactHTML div
   */
  public render(): React.ReactElement<Provider> {

    let item: any = (
      <Button
        buttonType={ButtonType.command}
        icon='ReplyAll'
        onClick={this.handleClick.bind(this) }>
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
   * @returns { string } Full HTML body with signature line
   */
  public addSignature(workItemHyperlink: string): string {
    return workItemHyperlink + '<br/><br/><br/>Created using VSTS Outlook add-in';
  }
  /**
   * Handles the click and displays a reply-all form
   * @private
   */
  private handleClick(): void {
    let props: IReplyAllButtonProps = this.props;
    if (Office.context.mailbox.diagnostics.hostName === 'OutlookIOS') {
      this.setState({ saving: true });
      Rest.autoReply('I have created the following bug:<br/><br/>' + this.addSignature(this.props.workItemHyperlink), (output: string) => {
        this.setState({ saving: false});
        props.dispatch(updateNotificationAction(NotificationType.Success, 'Reply Sent!'));
      });
    } else {
      Office.context.mailbox.item.displayReplyAllForm(this.addSignature(this.props.workItemHyperlink));
    }
  }
}
