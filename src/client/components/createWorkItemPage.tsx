import * as React from "react";
import { Description } from "./createWorkItem/description";
import { Title } from "./createWorkItem/title";
import { Save } from "./createWorkItem/save";
import { WorkItemDropdown } from "./createWorkitem/WorkItemDropdown";
import { Classification } from "./addConfig/classification";
import { Gear } from "./createWorkItem/gear";
import { Feedback } from "./shared/feedback";
import { Notification } from "./shared/notification";
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
export default class CreateWorkItem extends React.Component<ICreateWorkItemProps, {}> {
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
