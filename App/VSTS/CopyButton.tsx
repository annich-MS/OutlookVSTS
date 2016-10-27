/// <reference path="../../office.d.ts" />
import * as React from 'react';
import { Button, ButtonType } from 'office-ui-fabric-react';

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
    let id: string = 'Clipboard-Item';
    let existsTextarea: HTMLTextAreaElement = document.getElementById(id) as HTMLTextAreaElement;

    if (!existsTextarea) {
      let textarea: HTMLTextAreaElement = document.createElement('textarea');
      textarea.id = id;
      let style: any = {
        background: 'transparent',
        height: '1px',
        left: 0,
        padding: 0,
        position: 'fixed',
        top: 0,
        width: '1px',
      };
      Object.keys(style).forEach( key => {
          textarea.style.setProperty(key, style[key]);
      });

      document.querySelector('body').appendChild(textarea);
      existsTextarea = document.getElementById(id) as HTMLTextAreaElement;
    }

    existsTextarea.value = this.props.workItemHyperlink;
    existsTextarea.select();

    document.execCommand('copy');
  }
}
