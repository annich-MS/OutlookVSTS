import * as React from "react";
import { observer } from "mobx-react";
import { Dropdown, IDropdownOption } from "office-ui-fabric-react";

import IDropdownConfig from "../../stores/dropdownConfigStore";
import NavigationStore from "../../stores/navigationStore";

import { AppNotificationType } from "../../models/appNotification";

interface IConfigurableDropdownProps {
    dropdownConfig: IDropdownConfig;
    navigationStore: NavigationStore;
}

@observer
export default class ConfigurableDropdown extends React.Component<IConfigurableDropdownProps, {}> {

    public componentWillMount() {
        this.props.dropdownConfig.handleFailure = this.handleError.bind(this);
        this.props.dropdownConfig.populateIfNeeded();
    }

    public render(): JSX.Element {
        return (
            <Dropdown
                label={this.props.dropdownConfig.label}
                options={this.props.dropdownConfig.options}
                selectedKey={this.props.dropdownConfig.selected}
                onChanged={this.onChanged.bind(this)}
                disabled={this.props.dropdownConfig.isDisabled}
            />);
    }

    private handleError(error): void {
        this.props.navigationStore.updateNotification({ message: error.toString(), type: AppNotificationType.Error });
    }

    private onChanged(option: IDropdownOption): void {
        this.props.dropdownConfig.changeSelected(option);
    }
}
