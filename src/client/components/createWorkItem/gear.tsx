import * as React from "react";
import { IconButton } from "office-ui-fabric-react/lib/Button";
import NavigationStore from "../../stores/navigationStore";
import NavigationPage from "../../models/navigationPage";

/**
 * Represents the Gear Properties
 */
export interface IGearProps {
  navigationStore: NavigationStore;
}

/**
 * Renders the Gear Icon and the button underneath
 * @class { Gear }
 */
export class Gear extends React.Component<IGearProps, {}> {
  /**
   * Renders the Gear Icon and the button underneath
   */
  public render(): JSX.Element {
    return (
      <div style={{ float: "right" }}>
        <IconButton icon="Settings" title="Settings" onClick={this.handleGearClick.bind(this)} />
      </div>
    );
  }

  /**
   * Dispatches the action to change the pageVisibility value in the store
   * @ returns {void}
   */
  private handleGearClick(): void {
    this.props.navigationStore.clearNotification();
    this.props.navigationStore.navigate(NavigationPage.Settings);
  }
}
