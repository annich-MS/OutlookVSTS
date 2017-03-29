import * as React from "react";
import * as ReactDOM from "react-dom/server";
import { Provider } from "react-redux";
import { Label, Link } from "office-ui-fabric-react";

import { ItemHyperlink } from "./quickActions/itemHyperlink";
import { ReplyAllButton } from "./quickActions/replyAllButton";
import { CopyButton } from "./quickActions/copyButton";

import { Feedback } from "./shared/feedback";
import { Notification } from "./shared/notification";

import WorkItemStore from "../stores/workItemStore";
import NavigationStore from "../stores/navigationStore";

import { typeToString } from "../models/workItemType";

/**
 * Props for QuickActions Component
 */
interface IQuickActionProps {
  workItem: WorkItemStore;
  navigationStore: NavigationStore;
}

/**
 * Builds the formatted work item HTML element
 * Renders all Components
 */
export default class QuickActions extends React.Component<IQuickActionProps, {}> {
  /**
   * Builds the HTML element in the form <item type><item ID>: <item title>
   */
  public buildItemHyperlink(): string {
    return ReactDOM.renderToStaticMarkup(
      <Label className="WorkItemLink">
        <Link target="_blank" href={this.props.workItem.vstsInfo.vstsUrl}>
          {typeToString(this.props.workItem.type)} {this.props.workItem.vstsInfo.id}
        </Link>
        : {this.props.workItem.title}
      </Label>);
  }

  public buildTextOnly(): string {
    return `${typeToString(this.props.workItem.type)} ${this.props.workItem.vstsInfo.id}: ${this.props.workItem.title}`;
  }

  /**
   * Renders the ItemHyperlink, FollowButton, ReplyAllButton, and CopyButton Components
   */
  public render(): React.ReactElement<Provider> {
    let htmlString: string = this.buildItemHyperlink();
    let textString: string = this.buildTextOnly();
    return (
      <div>
        <Notification navigationStore={this.props.navigationStore} />
        <div className="ms-font-m-plus">Work item successfully created!</div>
        <br />
        <ItemHyperlink workItemHyperlink={htmlString} />
        <br />
        <br />
        <div className="ms-font-m-plus">Quick Actions:</div>
        <ReplyAllButton workItemHyperlink={htmlString} navigationStore={this.props.navigationStore} />
        <CopyButton workItemHyperlink={htmlString} textOnly={textString} />
        <br />
        <Feedback />
      </div>
    );
  }
}
