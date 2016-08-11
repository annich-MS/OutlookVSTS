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
    let buttonStyle: any = {
      background: 'rgb(255,255,255)',
      border: 'rgb(255,255,255)',
      color: 'rgb(0,0,0)',
      float: 'left',
      font: '15px arial, ms-segoe-ui',
    };
    return (
      <div>
        <button style={buttonStyle} onClick={this.handleClick.bind(this)}>
          <a className='ms-Icon ms-Icon--replyAll' />
          {'   '}Reply All with Work Item
        </button>
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
  private handleClick: () => void = () => {
    Office.context.mailbox.item.displayReplyAllForm(this.addSignature(this.props.workItemHyperlink));
  }
}
