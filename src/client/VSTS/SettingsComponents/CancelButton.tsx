import * as React from "react";
import { Button, ButtonType } from "office-ui-fabric-react";
import NavigationPage from "../../models/navigationPage";
import NavigationStore from "../../stores/navigationStore";

interface ISettingsProps {
  navigationStore: NavigationStore;
}

/**
 * renders the cancel button that redirects to CreateWorkItem
 */
export class CancelButton extends React.Component<ISettingsProps, any> {

  /**
   * Redirects to CreateWorkItem page
   */
  public Cancel(): void {
    this.props.navigationStore.navigate(NavigationPage.CreateWorkItem);
  }

  /**
   * Renders the cancel button
   */
  public render(): JSX.Element {
    return (
       <div style={{float: "right"}}>
          <Button buttonType={ButtonType.command} onClick={this.Cancel.bind(this)}>Cancel</Button>
      </div>
    );
  }
}
