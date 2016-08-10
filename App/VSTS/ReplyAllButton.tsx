/// <reference path="../../office.d.ts" />
import * as React from 'react';
import { Provider } from 'react-redux';

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
        <button onClick={this.handleClick.bind(this)} className='ms-Button'>
          <a className='ms-Icon ms-Icon--replyAll' />
          {'   '}Reply All with Work Item
        </button>
        <br/><br/>
      </div>
    );
  }
  
  public addSignature(workItemHyperlink: string): string {
    return workItemHyperlink + '<br/><br/><br/>Sent from VSTS Add-in for Outlook';
  }
  /**
   * Handles the click and displays a reply-all form
   * @private
   */
  private handleClick: () => void = () => {
    console.log(this.addSignature(this.props.workItemHyperlink));
    Office.context.mailbox.item.displayReplyAllForm(this.addSignature(this.props.workItemHyperlink));
  }
}
