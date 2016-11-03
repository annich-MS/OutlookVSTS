/// <reference path="../../office.d.ts" />
import * as React from 'react';
import { Button, ButtonType } from 'office-ui-fabric-react';
import * as Clipboard from 'clipboard-js';

/**
 * Props for CopyButton Component
 * @interface { ICopyButtonProps }
 */
interface ICopyButtonProps {
  /**
   * workItemHyperlink
   * @type { string }
   */
  workItemHyperlink: string;
  /** 
   * textOnly
   * @type {string}
   */
  textOnly: string;
}

/**
 * Renders a button that writes the item hyperlink HTML element to clipboard on-click
 * @class { CopyButton }
 */
export class CopyButton extends React.Component<ICopyButtonProps, {}> {

  /**
   * Renders the CopyButton Component and reads ICopyButtonProps
   * @returns { React.ReactElement } ReactHTML div 
   */
  public render(): React.ReactElement<{}> {

    return (
      <div>
        <Button
          buttonType={ButtonType.command}
          icon='Copy'
          onClick={this.handleClick}>
          Copy to Clipboard
        </Button>
      </div>
    );
  }

  /**
   * Handles the button click and fires a copy command
   * @private
   */
  private handleClick: () => void = () => {
    this.ieClipboard();
    /*
    switch (Office.context.mailbox.diagnostics.hostName) {
      case '':
        break;
      default:
        this.defaultClipboard();
    }
    */
  }

  private defaultClipboard(): void {
    Clipboard.copy({
      'text/plain': this.props.textOnly,
      'text/html': this.props.workItemHyperlink,
    });
  }

  private ieClipboard(): void {
    // select the email link anchor text
    let emailLink: Element = document.querySelector('.WorkItemLink');
    let range: Range = document.createRange();
    range.selectNode(emailLink);
    window.getSelection().addRange(range);

    try {
      // now that we've selected the anchor text, execute the copy command
      let successful: boolean = document.execCommand('copy');
      let msg: string = successful ? 'successful' : 'unsuccessful';
      console.log('Copy email command was ' + msg);
    } catch (err) {
      console.log('Oops, unable to copy');
    }
    // remove the selections - NOTE: Should use
    // removeRange(range) when it is supported  
    window.getSelection().removeAllRanges();
  }
}
