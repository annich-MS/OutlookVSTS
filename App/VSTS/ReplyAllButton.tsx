/// <reference path="../../office.d.ts" />
import * as React from 'react';
import { Provider } from 'react-redux';
import { Button, ButtonType } from 'office-ui-fabric-react';
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
export class ReplyAllButton extends React.Component<IReplyAllButtonProps, {}> {
  /**
   * Renders the ReplyAllButton Component and reads IReplyAllButtonProps
   * @returns { React.ReactElement } ReactHTML div
   */
  public render(): React.ReactElement<Provider> {

    return (
      <div>
        <Button
          buttonType={ButtonType.command}
          icon='ReplyAll'
          onClick={this.handleClick.bind(this)}>
          Reply All with Work Item
        </Button>
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
      Rest.log('ios');
      Office.context.mailbox.getCallbackTokenAsync((asyncResult: Office.AsyncResult) => {
        Rest.log('got callback token: ' + JSON.stringify(asyncResult));
        if (asyncResult.error) {
          props.dispatch(updateNotificationAction(NotificationType.Error, 'Reply failed due to ' + asyncResult.error));
        } else {
          Rest.log('Sending Message');
          let settings: any = {
            contentType: 'application/json',
            data: JSON.stringify({
              'Comment': 'I have created the following bug:<br/><br/>' +  this.addSignature(this.props.workItemHyperlink),
            }),
            error: function (a: any, b: any, c: any): void {
              Rest.log('Error' + b);
              props.dispatch(updateNotificationAction(NotificationType.Error, 'Reply failed due to ' + b));
            },
            headers: {
              'Authorization': 'Bearer ' + asyncResult.value,
            },
            success: function (): void {
              Rest.log('Success!');
              props.dispatch(updateNotificationAction(NotificationType.Success, 'Done!'));
            },
            url: 'https://outlook.office365.com/api/v2.0/me/messages/' + Office.context.mailbox.item.itemId + '/replyAll',
          };
          $.post(settings).done(() => {
            Rest.log('Success!');
            props.dispatch(updateNotificationAction(NotificationType.Error, 'Reply failed due to ' + asyncResult.error));
          }).fail((jqXHR, status, errorThrown) => {
            Rest.log('Error' + status);
            props.dispatch(updateNotificationAction(NotificationType.Success, 'Reply succeded due to ' + errorThrown));
          });
        }
      });
    } else {
      Office.context.mailbox.item.displayReplyAllForm(this.addSignature(this.props.workItemHyperlink));
    }
  }
}
