import * as React from "react";
import { CommandBar, DetailsList, IContextualMenuItem, Selection, SelectionMode, IColumn, ConstrainMode } from "office-ui-fabric-react";
import NavigationStore from "../../stores/navigationStore";
import NavigationPage from "../../models/navigationPage";
import IVSTSConfig from "../../models/vstsConfig";
import VSTSConfigStore from "../../stores/vstsConfigStore";

interface IConfigDisplayProps {
    navigationStore: NavigationStore;
    vstsConfig: VSTSConfigStore;
}

interface IConfigDisplayState {
    selected: string;
}

export default class ConfigDisplay extends React.Component<IConfigDisplayProps, IConfigDisplayState> {

    private readonly items: IContextualMenuItem[] = [
        {
            icon: "Add",
            key: "addConfig",
            name: "Add",
            onClick: this.addConfig.bind(this),
        },
        {
            icon: "Delete",
            key: "deleteConfig",
            name: "Remove",
            onClick: this.removeConfig.bind(this),
        },
    ];

    private readonly columns: IColumn[] = [
        {
            fieldName: "name",
            key: "name",
            minWidth: 1,
            name: "Config Name",
        }
    ];

    private selection: Selection;

    public constructor() {
        super();
        this.selection = new Selection({
            onSelectionChanged: () => this.setState({ selected: this.getSelectionName() }),
        });
    }

    public render(): React.ReactElement<any> {
        return (
            <div style={{ overflow: "hidden" }} >
                <CommandBar items={this.items} />
                <DetailsList
                    items={this.props.vstsConfig.configs}
                    constrainMode={ConstrainMode.unconstrained}
                    columns={this.columns}
                    selection={this.selection}
                    selectionMode={SelectionMode.single} />
            </div>);
    }

    private addConfig(): void {
        this.props.navigationStore.navigate(NavigationPage.AddConfig);
    }

    private removeConfig(): void {
        if (this.state.selected !== "") {
            this.props.vstsConfig.removeConfig(this.state.selected);
            this.forceUpdate();
            if (this.props.vstsConfig.configs.length === 0) {
                this.props.navigationStore.navigate(NavigationPage.AddConfig);
            }
        }
    }

    private getSelectionName() {
        if (this.selection.getSelectedCount() === 0) {
            return "";
        }
        return (this.selection.getSelection()[0] as IVSTSConfig).name;
    }
}
