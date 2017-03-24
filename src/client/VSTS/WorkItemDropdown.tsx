import * as React from "react";
import { Dropdown, IDropdownOption } from "office-ui-fabric-react";
import { observer } from "mobx-react";
import { typeFromString, typeToString } from "../models/workItemType";
import WorkItemStore from "../stores/workItemStore";

/**
 * Represents the WorkItemType Properties
 */
export interface IWorkItemTypeDropdownProps {
    workItem: WorkItemStore;
}

@observer
export class WorkItemDropdown extends React.Component<IWorkItemTypeDropdownProps, {}> {
    /**
     * Dipatches an action to update the value of workItemType in the store to the selected value
     */
    public handleTypeChange(option: IDropdownOption): void {
        this.props.workItem.setType(typeFromString(option.text));
    }
    /**
     * Renders the workItemType Dropdown using React-Select
     */
    public render(): JSX.Element {

        let types: any = [
            { key: "Bug", text: "Bug" },
            { key: "Task", text: "Task" },
            { key: "User Story", text: "User Story" },
        ];

        return (
            <div>
                <br />
                <Dropdown
                    label="Work Item Type"
                    options={types}
                    selectedKey={typeToString(this.props.workItem.type)}
                    onChanged={this.handleTypeChange.bind(this)} />
            </div>
        );
    }
}
