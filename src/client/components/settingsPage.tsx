import * as React from "react";

import { Notification } from "./shared/notification";
import { LogoutButton } from "./settings/logoutButton";
import NavigationStore from "../stores/navigationStore";
import ConfigDisplay from "./settings/configDisplay";
import VSTSConfigStore from "../stores/vstsConfigStore";
import CancelButton from "./shared/cancelButton";
import PageTitle from "./shared/pageTitle";
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

    return (
      <div>
        <Notification navigationStore={this.props.navigationStore} />
        <CancelButton navigationStore={this.props.navigationStore} backTarget={NavigationPage.CreateWorkItem} />
        <br />
        <PageTitle>Configurations</PageTitle>
        <ConfigDisplay navigationStore={this.props.navigationStore} vstsConfig={this.props.vstsConfig} />
        <br />
        <LogoutButton navigationStore={this.props.navigationStore} />
      </div>
    );
  }
}
