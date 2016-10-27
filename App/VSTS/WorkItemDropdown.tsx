import * as React from 'react';
import { Provider, connect } from 'react-redux';
import { updateWorkItemType } from '../Redux/WorkItemActions';
import { Dropdown } from 'office-ui-fabric-react';

/**
 * Represents the WorkItemType Properties
 * @interface IWorkItemTypeDropdownProps
 */
export interface IWorkItemTypeDropdownProps {
    /**
     * dispatch to map dispatch to props
     * @type {any}
     */
    dispatch?: any;
    /**
     * represents the type of work item the user selects
     * @type {string}
     */
    workItemType?: string;
}

/**
 * Renders the dropdown to select the workItemType using React-Select
 * @class { WorkItemDropdown }
 */
function mapStateToProps(state: any): IWorkItemTypeDropdownProps {
    return { workItemType: state.workItem.workItemType };
}

@connect(mapStateToProps)
export class WorkItemDropdown extends React.Component<IWorkItemTypeDropdownProps, {}> {
    /**
     * Dipatches an action to update the value of workItemType in the store to the selected value
     * @returns {void}
     * @param {any} option
     */
    public handleTypeChange(option: any): void {
        let type: string;
        if (option.text) {
            type = option.text;
        } else {
            type = option;
        }
        this.props.dispatch(updateWorkItemType(type));
    }
    /**
     * Renders the workItemType Dropdown using React-Select
     */
    public render(): React.ReactElement<Provider> {

        let types: any = [
            { key: 'Bug', text: 'Bug' },
            { key: 'Task', text: 'Task' },
            { key: 'User Story', text: 'User Story' },
        ];

        return (
            <div>
                <br/>
                <Dropdown
                    label={'Work Item Type'}
                    options={types}
                    selectedKey={this.props.workItemType}
                    onChanged={this.handleTypeChange.bind(this)} />
            </div>
        );
    }
}
