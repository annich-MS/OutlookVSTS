import * as React from "react";
import { observer } from "mobx-react";
import { RoamingSettings } from "../RoamingSettings";
import { Button, ButtonType } from "office-ui-fabric-react";
import APTCache from "../../stores/aptCache";
import NavigationStore from "../../stores/navigationStore";
import NavigationPage from "../../models/navigationPage";
import APTPopulateStage from "../../models/aptPopulateStage";

interface ISettingsProps {
  cache: APTCache;
  navigationStore: NavigationStore;
}

/**
 * Smart component
 * Renders area path dropdowns and save button
 * @class {Settings} 
 */
@observer
export class SaveDefaultsButton extends React.Component<ISettingsProps, any> {

  /**
   * saves current selected settings to Office Roaming Settings
   * updates page state to Create Work Item page
   * @returns {void}
   */
  public saveDefaults(): void {
    let rs: RoamingSettings = RoamingSettings.GetInstance();
    rs.updateFromCache(this.props.cache);
    rs.save();
    this.props.navigationStore.navigate(NavigationPage.CreateWorkItem);
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
          onClick={this.saveDefaults.bind(this)}
          disabled={this.props.cache.populateStage < APTPopulateStage.PostPopulate}>
          Save and continue
          </Button>
      </div>
    );
  }
}
