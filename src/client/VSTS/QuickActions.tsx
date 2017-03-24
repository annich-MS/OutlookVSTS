import * as React from "react";
import { Provider } from "react-redux";
import { ItemHyperlink } from "./ItemHyperlink";
import { ReplyAllButton } from "./ReplyAllButton";
import { CopyButton } from "./CopyButton";
import { Feedback } from "./SimpleComponents/Feedback";
import * as ReactDOM from "react-dom/server";
import { Label, Link } from "office-ui-fabric-react";
import { Notification } from "./SimpleComponents/Notification";
import WorkItemStore from "../stores/workItemStore";
import NavigationStore from "../stores/navigationStore";

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
export class QuickActions extends React.Component<IQuickActionProps, {}> {
  /**
   * Builds the HTML element in the form <item type><item ID>: <item title>
   */
  public buildItemHyperlink(): string {
    return ReactDOM.renderToStaticMarkup(
      <Label className="WorkItemLink">
        <Link target="_blank" href={this.props.workItem.vstsInfo.vstsUrl}>
          {this.props.workItem.type} {this.props.workItem.vstsInfo.id}
        </Link>
        : {this.props.workItem.title}
      </Label>);
  }

  public buildTextOnly(): string {
    return `${this.props.workItem.type} ${this.props.workItem.vstsInfo.id}: ${this.props.workItem.title}`;
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
