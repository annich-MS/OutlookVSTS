import * as React from "react";
import { TextField } from "office-ui-fabric-react";
import { observer } from "mobx-react";

import WorkItemStore from "../../stores/workItemStore";

/**
 * Represents the Title Properties
 */
interface ITitleProps {
  workItem: WorkItemStore;
}

@observer
export default class ConfigName extends React.Component<ITitleProps, {}> {
  /**
   * Dipatches the action to change the value of title in the store 
   */
  public handleChangeTitle(value: string): void {
    this.props.workItem.setTitle(value);
  }
  /**
   * Rendersthe Title heading and the Title textbox
   */
  public render(): JSX.Element {
    return (
      <div>
        <TextField
          label="Configuration name"
          onChanged={this.handleChangeTitle.bind(this)}
          value={this.props.workItem.title} />
      </div>
    );
  }
}