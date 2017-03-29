import * as React from "react";
import { Notification } from "../SimpleComponents/Notification";
import { Classification } from "./Classification";
import { SaveDefaultsButton } from "./SaveDefaultsButton";
import { CancelButton } from "./CancelButton";
import { LogoutButton } from "./LogoutButton";
import NavigationStore from "../../stores/navigationStore";
import APTCache from "../../stores/aptCache";

interface ISettingsProps {
  cache: APTCache;
  navigationStore: NavigationStore;
}

/**
 * Smart component
 * Renders area path dropdowns and save button
 * @class {Settings} 
 */
export class Settings extends React.Component<ISettingsProps, any> {
  /**
   * Renders the area path dropdowns and save button
   */
  public render(): JSX.Element {
    let textStyle: string = "ms-font-m-plus";

    return (
      <div>
        <Notification navigationStore={this.props.navigationStore} />
        <div>
          <p className={textStyle}> Welcome {Office.context.mailbox.userProfile.displayName}!</p>
          <p />
          <p className={textStyle}> Take a moment to configure your default settings for work item creation.</p>
        </div>
        <div>
          <Classification cache={this.props.cache} navigationStore={this.props.navigationStore}/>
        </div>
        <div>
          <SaveDefaultsButton cache={this.props.cache} navigationStore={this.props.navigationStore} />
          <CancelButton navigationStore={this.props.navigationStore} />
        </div>
        <br />
        <div>
          <LogoutButton aptCache={this.props.cache} navigationStore={this.props.navigationStore} />
        </div>
      </div>
    );
  }
}
