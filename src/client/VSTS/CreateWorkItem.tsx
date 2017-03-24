import * as React from "react";
import { Description } from "./Description";
import { Title } from "./Title";
import { Save } from "./Save";
import { WorkItemDropdown } from "./WorkItemDropdown";
import { Classification } from "./SettingsComponents/Classification";
import { Gear } from "./Gear";
import { Feedback } from "./SimpleComponents/Feedback";
import { Notification } from "./SimpleComponents/Notification";
import APTCache from "../stores/aptCache";
import NavigationStore from "../stores/navigationStore";
import WorkItemStore from "../stores/workItemStore";

interface ICreateWorkItemProps {
  cache: APTCache;
  navigationStore: NavigationStore;
  workItem: WorkItemStore;
}

/**
 * Renders all components of the Create page
 */
export class CreateWorkItem extends React.Component<ICreateWorkItemProps, {}> {
  /**
   * Renders the div that contains all the components of the Create page
   */
  public render(): React.ReactElement<{}> {
    return (
      <div>
        <Notification navigationStore={this.props.navigationStore} />
        <Gear navigationStore={this.props.navigationStore} />
        <WorkItemDropdown workItem={this.props.workItem} />
        <Title workItem={this.props.workItem} />
        <Description workItemStore={this.props.workItem} />
        <Classification cache={this.props.cache} navigationStore={this.props.navigationStore} />
        <Save cache={this.props.cache} navigationStore={this.props.navigationStore} workItem={this.props.workItem} />
        <Feedback/>
      </div>
    );
  }
}
