import * as React from "react";

import { Notification } from "./shared/notification";
import CancelButton from "./shared/cancelButton";

import { Classification } from "./addConfig/classification";
import SaveConfigButton from "./addConfig/saveConfigButton";

import NavigationStore from "../stores/navigationStore";
import APTCache from "../stores/aptCache";
import VSTSConfigStore from "../stores/vstsConfigStore";
import ConfigName from "./addConfig/configName";
import { Feedback } from "./shared/feedback";

import NavigationPage from "../models/navigationPage";

interface ISettingsProps {
    cache: APTCache;
    navigationStore: NavigationStore;
    vstsConfig: VSTSConfigStore;
}

/**
 * Smart component
 * Renders area path dropdowns and save button
 * @class {Settings} 
 */
export default class Settings extends React.Component<ISettingsProps, any> {
    /**
     * Renders the area path dropdowns and save button
     */
    public render(): JSX.Element {
        let textStyle: string = "ms-font-m-plus";

        return (
            <div>
                <Notification navigationStore={this.props.navigationStore} />
                <div>
                    <p className={textStyle}> Create a bug creation configuration.</p>
                </div>
                <div>
                    <ConfigName cache={this.props.cache}/>
                    <Classification cache={this.props.cache} navigationStore={this.props.navigationStore} />
                </div>
                <div>
                    <SaveConfigButton cache={this.props.cache} navigationStore={this.props.navigationStore} vstsConfig={this.props.vstsConfig} />
                    <CancelButton navigationStore={this.props.navigationStore} backTarget={NavigationPage.Settings} />
                </div>
                <br />
                <div>
                    <Feedback />
                </div>
            </div>
        );
    }
}
