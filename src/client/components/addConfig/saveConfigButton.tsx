import * as React from "react";
import { observer } from "mobx-react";
import { Button, ButtonType } from "office-ui-fabric-react";

import APTCache from "../../stores/aptCache";
import NavigationStore from "../../stores/navigationStore";

import VSTSConfigStore from "../../stores/vstsConfigStore";
import IVSTSConfig from "../../models/vstsConfig";
import { AppNotificationType } from "../../models/appNotification";

interface ISaveConfigButtonProps {
  cache: APTCache;
  navigationStore: NavigationStore;
  vstsConfig: VSTSConfigStore;
}

/**
 * Smart component
 * Renders area path dropdowns and save button
 * @class {Settings} 
 */
@observer
export default class SaveConfigButton extends React.Component<ISaveConfigButtonProps, any> {

  /**
   * saves current selected settings to Office Roaming Settings
   * updates page state to Create Work Item page
   * @returns {void}
   */
  public async save(): Promise<void> {
    let unique: boolean = this.props.vstsConfig.configs.filter((config: IVSTSConfig) => { return config.name === this.props.cache.name; }).length === 0;
    if (!unique) {
      this.props.navigationStore.updateNotification({
        message: `cannot create config with name "${this.props.cache.name}" as one already exists.`,
        type: AppNotificationType.Warning,
      });
      return;
    } else {
      let config: IVSTSConfig = {
        account: this.props.cache.account,
        name: this.props.cache.name,
        project: this.props.cache.project,
        team: this.props.cache.team,
      };
      this.props.vstsConfig.addConfig(config);
      this.props.navigationStore.navigateBack();
    }
    return;
  }

  /**
   * Renders the area path dropdowns and save button
   */
  public render(): JSX.Element {
    return (
      <div style={{ float: "left" }}>
        <Button
          buttonType={ButtonType.command}
          icon="Save"
          onClick={this.save.bind(this)}>
          Save Configuration
          </Button>
      </div>
    );
  }
}
