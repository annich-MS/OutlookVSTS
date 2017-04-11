import * as React from "react";
import { List } from "office-ui-fabric-react/lib/List";
import { PrimaryButton } from "office-ui-fabric-react/lib/Button";
import NavigationStore from "../../stores/navigationStore";
import NavigationPage from "../../models/navigationPage";
import IVSTSConfig from "../../models/vstsConfig";
import VSTSConfigStore from "../../stores/vstsConfigStore";
import ConfigOption from "./configOption";

interface IConfigDisplayProps {
    navigationStore: NavigationStore;
    vstsConfig: VSTSConfigStore;
}

export default class ConfigDisplay extends React.Component<IConfigDisplayProps, {}> {

    public render(): React.ReactElement<any> {
        return (
            <div style={{ overflow: "hidden" }} >
                <List items={this.props.vstsConfig.configs} onRenderCell={this.renderCell.bind(this)} />
                <br />
                <div style={{ margin: "auto", textAlign: "center", width: "75%" }}>
                    <PrimaryButton icon="Add" onClick={this.addConfig.bind(this)}>New</PrimaryButton>
                </div>
            </div>);
    }

    private renderCell(item: IVSTSConfig, index: number): JSX.Element {
        return (<ConfigOption item={item} removeConfig={this.removeConfig.bind(this)} />);
    }


    private removeConfig(name: string): void {
        this.props.vstsConfig.removeConfig(name);
        if (this.props.vstsConfig.configs.length === 0) {
            this.props.navigationStore.navigate(NavigationPage.AddConfig);
        } else {
            this.forceUpdate();
        }
    }

    private addConfig(): void {
        this.props.navigationStore.navigate(NavigationPage.AddConfig);
    }

}
