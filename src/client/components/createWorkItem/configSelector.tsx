import * as React from "react";
import { observer } from "mobx-react";
import { Dropdown, IDropdownOption } from "office-ui-fabric-react";

import IVSTSConfig from "../../models/vstsConfig";

import VSTSConfigStore from "../../stores/vstsConfigStore";

interface IConfigSelectorProps {
    configStore: VSTSConfigStore;
}

@observer
export default class ConfigSelector extends React.Component<IConfigSelectorProps, {}> {

    public render(): JSX.Element {
        return (
            <Dropdown
                label="Configuration"
                options={this.props.configStore.configs.map((config: IVSTSConfig) => { return { key: config.name, text: config.name }; })}
                selectedKey={this.props.configStore.selected}
                onChanged={this.onChanged.bind(this)}
            />);
    }

    private onChanged(option: IDropdownOption): void {
        this.props.configStore.setSelected(option.text);
    }
}