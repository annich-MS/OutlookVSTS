import * as React from "react";
import { Description } from "./createWorkItem/description";
import { Title } from "./createWorkItem/title";
import Save from "./createWorkItem/save";
import { WorkItemDropdown } from "./createWorkitem/WorkItemDropdown";
import { Gear } from "./createWorkItem/gear";
import { Feedback } from "./shared/feedback";
import { Notification } from "./shared/notification";
import APTCache from "../stores/aptCache";
import NavigationStore from "../stores/navigationStore";
import WorkItemStore from "../stores/workItemStore";
import ConfigSelector from "./createWorkItem/configSelector";
import VSTSConfigStore from "../stores/vstsConfigStore";

interface ICreateWorkItemProps {
  navigationStore: NavigationStore;
  workItem: WorkItemStore;
  vstsConfigStore: VSTSConfigStore;
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
        <ConfigSelector configStore={this.props.vstsConfigStore} />
        <Save navigationStore={this.props.navigationStore} workItem={this.props.workItem} vstsConfig={this.props.vstsConfigStore} />
        <Feedback/>
      </div>
    );
  }
}
