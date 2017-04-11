import * as React from "react";
import { observer } from "mobx-react";
import { Checkbox } from "office-ui-fabric-react/lib/Checkbox";
import * as Editable from "react-contenteditable";
import WorkItemStore from "../../stores/workItemStore";
import Constants from "../../models/constants";

/**
 * Represents the Description Properties
 */
export interface IDescriptionProps {
  workItemStore: WorkItemStore;
}

@observer
export class Description extends React.Component<IDescriptionProps, {}> {

  private CHECKBOX_LABEL: string = "Add e-mail as attachment";

  /**
   * Dispatches the action to change the description value in the store
   */
  public handleChangeDescription(event: any): void {
    this.props.workItemStore.setDescription(event.target.value);
  }

  /**
   * Dispatches the action to update the addAsAttachment and description values in the store
   */
  public handleChangeAddAsAttachment(event: any, isChecked: boolean): void {
    this.props.workItemStore.toggleAttachEmail();
  }

  /**
   * Renders the Description heading, the Add Email as Attachment checkbox, and the Description textbox
   */
  public render(): JSX.Element {

    let checkbox: any = null;

    if (Office.context.mailbox.diagnostics.hostName !== Constants.IOS_HOST_NAME) {
      checkbox = <Checkbox label={this.CHECKBOX_LABEL} onChange={this.handleChangeAddAsAttachment.bind(this)} defaultChecked={true} />;
    }
    return (
      <div>
        {checkbox}
        <Editable
          className={"editable"}
          disabled={false}
          html={this.props.workItemStore.description}
          onChange={this.handleChangeDescription.bind(this)} />
      </div>
    );
  }
}

