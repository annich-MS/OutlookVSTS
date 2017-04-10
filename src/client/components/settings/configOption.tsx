import * as React from "react";

import IVSTSConfig from "../../models/vstsConfig";

interface IConfigOptionProps {
    item: IVSTSConfig;
    removeConfig: (name: string) => void;
}

export default class ConfigOption extends React.Component<IConfigOptionProps, {}> {

    public render(): JSX.Element {
        let item: IVSTSConfig = this.props.item;
        return (
            <div className="ms-ListItem ms-ListItem--document">
                <div style={{ paddingLeft: "10px" }} >
                    <span className="ms-ListItem-primaryText" >{item.name}</span>
                    <div className="ms-ListItem-secondaryText" style={{ clear: "left" }} >{item.account}|{item.project}|{item.team}</div>
                    <div className="ms-ListItem-actions">
                        <div className="ms-ListItem-action">
                            <i className="ms-Icon ms-Icon--Delete" onClick={() => this.props.removeConfig(item.name)} />
                        </div>
                    </div>
                </div>
            </div>);
    }
}
