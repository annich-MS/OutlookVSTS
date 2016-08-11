/// <reference path="../../office.d.ts" />
import * as React from 'react';

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
    let buttonStyle: any = {
      background: 'rgb(255,255,255)',
      border: 'rgb(255,255,255)',
      color: 'rgb(0,0,0)',
      float: 'left',
      font: '15px arial, ms-segoe-ui',
    };
    document.addEventListener('copy', this.setClipboardData);
    return (
      <div>
        <button style={buttonStyle} onClick={this.handleClick}>
          <a className='ms-Icon ms-Icon--copy'/>
          {'   '}Copy to Clipboard
        </button>
      </div>
    );
  }

  /**
   * Handles the button click and fires a copy command
   * @private
   */
  private handleClick: () => void = () => {
    document.execCommand('copy');
  }

  /**
   * Writes the formatted HTML element to the clipboard
   * @private
   * @param { any } e 
   */
  private setClipboardData: (e: any) => void = (e) => {
    console.log('got to handler');
    e.clipboardData.setData('text/html', this.props.workItemHyperlink);
    e.preventDefault(); // we want our data, not data from any selection, to be written to the clipboard
  }
 }
