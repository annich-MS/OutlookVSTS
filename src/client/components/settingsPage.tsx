import * as React from "react";
import { Notification } from "./shared/notification";
import { CancelButton } from "./shared/cancelButton";
import { LogoutButton } from "./settings/logoutButton";
import NavigationStore from "../stores/navigationStore";
import ConfigDisplay from "./settings/configDisplay";
import VSTSConfigStore from "../stores/vstsConfigStore";
import NavigationPage from "../models/navigationPage";

interface ISettingsProps {
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
          <p style={{textAlign: "center"}} className={textStyle}> Configurations </p>
        </div>
        <div>
          <ConfigDisplay navigationStore={this.props.navigationStore} vstsConfig={this.props.vstsConfig} />
          <CancelButton navigationStore={this.props.navigationStore} backTarget={NavigationPage.CreateWorkItem} />
        </div>
        <br />
        <div>
          <LogoutButton navigationStore={this.props.navigationStore} />
        </div>
      </div>
    );
  }
}
