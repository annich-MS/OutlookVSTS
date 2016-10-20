/// <reference path="../../office.d.ts" />
import * as React from 'react';
import { Provider } from 'react-redux';
import { ItemHyperlink } from  './ItemHyperlink';
// import { FollowButton } from './FollowButton';
import { ReplyAllButton } from './ReplyAllButton';
import { CopyButton } from './CopyButton';
import { IWorkItem } from '../Redux/WorkItemReducer';
import { Feedback } from './SimpleComponents/Feedback';
import { connect } from 'react-redux';
import * as ReactDOM from 'react-dom/server';

/**
 * Props for QuickActions Component
 * @interface { IQuickActionProps }
 */
interface IQuickActionProps {
  /**
   * Work item information
   * @type { IWorkItem }
   */
  workItem?: IWorkItem;
}

/**
 * Mapping state from store to component props
 * @returns { IQuickActionProps } Props for QuickActions Component
 */
function mapStateToProps(state: any): IQuickActionProps {
  return {
    workItem: state.workItem,
  };
}

/**
 * Builds the formatted work item HTML element
 * Renders all Components
 * @returns { React.ReactElement } ReactHTML div
 */
@connect(mapStateToProps)
export class QuickActions extends React.Component<IQuickActionProps, {}> {
  /**
   * Builds the HTML element in the form <item type><item ID>: <item title>
   * @returns { string }
   */
  public buildItemHyperlink(): string {
    return ReactDOM.renderToStaticMarkup(
      <label>
        <a target='_blank'
          href={this.props.workItem.VSTShtmlLink}
          className='15px arial, ms-segoe-ui'>
          {this.props.workItem.workItemType} {this.props.workItem.id}
        </a>
        <a className='15px arial, ms-segoe-ui'>: {this.props.workItem.title}</a>
      </label>);
  }

  /**
   * Renders the ItemHyperlink, FollowButton, ReplyAllButton, and CopyButton Components
   * @returns { React.ReactElement } ReactHTML div
   */
  public render(): React.ReactElement<Provider> {
    let headerStyle: any = {
      font: '16px arial, ms-segoe-ui',
      'padding-bottom': '20px',
    };
    let htmlString: string = this.buildItemHyperlink();
    return(
      <div>
        <div style={headerStyle}>Work item successfully created!</div>
        <ItemHyperlink workItemHyperlink={htmlString}/>
        <div style={headerStyle}>Quick Actions:</div>
        <ReplyAllButton workItemHyperlink={htmlString}/>
        <br/>
        <CopyButton workItemHyperlink={this.props.workItem.VSTShtmlLink}/>
        <Feedback />
      </div>
    );
  }
}
