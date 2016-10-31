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
    Clipboard.copy({
      'text/plain': this.props.textOnly,
      'text/html': this.props.workItemHyperlink
    });
  }
}
