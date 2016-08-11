/// <reference path="../../office.d.ts" />
import * as React from 'react';

/**
 * Props for ItemHyperlink Component
 * @interface { IItemHyperlinkProps }
 */
interface IItemHyperlinkProps {
  /**
   * workItemHyperlink
   * @type { string }
   */
  workItemHyperlink: string;
}

/**
 * Renders a ReactHTML div 
 * @class { ItemHyperlink }
 */
export class ItemHyperlink extends React.Component<IItemHyperlinkProps, {}> {
  /**
   * Renders the ItemHyperlink Component and reads IItemHyperlinkProps
   * @returns { React.ReactElement } ReactHTML div
   */
  public render(): React.ReactElement<any> {
    let elementStyle: any = {
      'padding-bottom': '40px',
    };
    return(
      <div style={elementStyle}>
        <td dangerouslySetInnerHTML={{__html: this.props.workItemHyperlink}} />
      </div>
    );
  }
 }
