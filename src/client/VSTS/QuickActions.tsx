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
import { Label, Link } from 'office-ui-fabric-react';
import { Notification } from './SimpleComponents/Notification';

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

  dispatch?: any;
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
      <Label className='WorkItemLink'>
        <Link target='_blank' href={this.props.workItem.VSTShtmlLink}>
          {this.props.workItem.workItemType} {this.props.workItem.id}
        </Link>
        : {this.props.workItem.title}
      </Label>);
  }

  public buildTextOnly(): string {
    return this.props.workItem.workItemType + ' ' + this.props.workItem.id + ': ' + this.props.workItem.title;
  }

  /**
   * Renders the ItemHyperlink, FollowButton, ReplyAllButton, and CopyButton Components
   * @returns { React.ReactElement } ReactHTML div
   */
  public render(): React.ReactElement<Provider> {
    let htmlString: string = this.buildItemHyperlink();
    let textString: string = this.buildTextOnly();
    return(
      <div>
        <Notification />
        <div className='ms-font-m-plus'>Work item successfully created!</div>
        <br />
        <ItemHyperlink workItemHyperlink={htmlString}/>
        <br />
        <br />
        <div className='ms-font-m-plus'>Quick Actions:</div>
        <ReplyAllButton workItemHyperlink={htmlString} dispatch={this.props.dispatch} />
        <CopyButton workItemHyperlink={htmlString} textOnly={textString} />
        <br />
        <Feedback />
      </div>
    );
  }
}
